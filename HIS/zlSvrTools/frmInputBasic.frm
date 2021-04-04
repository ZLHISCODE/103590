VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmInputBasic 
   Caption         =   "基本字词管理"
   ClientHeight    =   5505
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8655
   Icon            =   "frmInputBasic.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   8655
   StartUpPosition =   1  '所有者中心
   Begin VSFlex8Ctl.VSFlexGrid msf 
      Height          =   2595
      Left            =   2460
      TabIndex        =   4
      Top             =   870
      Width           =   4620
      _cx             =   8149
      _cy             =   4577
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483636
      FocusRect       =   2
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   5
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   240
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
      WallPaper       =   "frmInputBasic.frx":1CFA
      WallPaperAlignment=   8
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.TreeView tvw 
      Height          =   1995
      Left            =   165
      TabIndex        =   3
      Top             =   825
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   3519
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   840
      Top             =   4155
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputBasic.frx":3EC58
            Key             =   "Item"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   5145
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmInputBasic.frx":3F532
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10213
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
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   1244
      BandCount       =   2
      _CBWidth        =   8655
      _CBHeight       =   705
      _Version        =   "6.7.8988"
      Child1          =   "tbrThis"
      MinHeight1      =   645
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Caption2        =   "类型"
      Child2          =   "cbo"
      MinWidth2       =   1500
      MinHeight2      =   300
      Width2          =   390
      NewRow2         =   0   'False
      Begin VB.ComboBox cbo 
         Height          =   300
         Left            =   7065
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   195
         Width           =   1500
      End
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsHotMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   11
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "预览"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "增加"
               Key             =   "增加"
               Object.ToolTipText     =   "增加"
               Object.Tag             =   "增加"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "修改"
               Object.ToolTipText     =   "修改"
               Object.Tag             =   "修改"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "删除"
               Object.ToolTipText     =   "删除"
               Object.Tag             =   "删除"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_3"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查找"
               Key             =   "查找"
               Object.ToolTipText     =   "查找"
               Object.Tag             =   "查找"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   8
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   7755
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputBasic.frx":3FDC6
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputBasic.frx":3FFE6
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputBasic.frx":40206
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputBasic.frx":40420
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputBasic.frx":4063A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputBasic.frx":40854
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputBasic.frx":40A6E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputBasic.frx":40C8E
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   7155
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputBasic.frx":40EAE
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputBasic.frx":410CE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputBasic.frx":412EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputBasic.frx":41508
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputBasic.frx":41728
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputBasic.frx":41948
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputBasic.frx":41B62
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputBasic.frx":41D82
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   825
      Top             =   3555
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
            Picture         =   "frmInputBasic.frx":41FA2
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInputBasic.frx":4253C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgY 
      Height          =   3435
      Left            =   2220
      MousePointer    =   9  'Size W E
      Top             =   855
      Width           =   45
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePrintView 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileOutExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "增加(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditAddCopy 
         Caption         =   "复制增加(&N)"
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除(&D)"
         Shortcut        =   {DEL}
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
      Begin VB.Menu mnuViewFind 
         Caption         =   "查找(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuView_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "帮助主题(&H)"
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
      Begin VB.Menu h1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmInputBasic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mstrArray(1 To 26) As String

Private mblnStartUp As Boolean
Private mlngLoop As Long
Private mstrLastKey As String
Private mlngCboIndex As Long

Private Sub SavePostion(ByVal bytMode As Byte, ByRef strName As String, ByRef strCode As String)
    '--------------------------------------------------------------------------------------
    '功能：保存表格的当前位置
    '--------------------------------------------------------------------------------------
    Dim lngLoop As Long
    
    If msf.TextMatrix(msf.Row, msf.Col) = "" Then Exit Sub
    
    If bytMode = 1 Then
        If tvw.SelectedItem Is Nothing Then Exit Sub
        
        If Left(tvw.SelectedItem.Key, 1) = "P" Then
            '1.拼音
            If msf.MergeRow(msf.Row) = False Then
                strName = msf.TextMatrix(msf.Row, msf.Col)
                
                '向上找拼音
                For lngLoop = msf.Row - 1 To 1 Step -1
                    If msf.MergeRow(lngLoop) Then
                        strCode = msf.TextMatrix(lngLoop, 0)
                        Exit For
                    End If
                Next
            End If
        Else
            '2.五笔
            strName = msf.TextMatrix(msf.Row, 0)
            strCode = msf.TextMatrix(msf.Row, msf.Col)
        End If
    Else
        strName = msf.TextMatrix(msf.Row, 0)
        strCode = msf.TextMatrix(msf.Row, 1)
    End If
End Sub

Private Sub LocationPostion(ByVal strName As String, ByVal strCode As String)
    Dim lngLoop As Long
    Dim lngCount As Long
    Dim lngCol As Long
    
    On Error Resume Next
    
    If Val(Left(cbo.Text, 1)) = 1 Then
        '基本字，分拼音和五笔两种
        
        If tvw.SelectedItem Is Nothing Then Exit Sub
                
        If Left(tvw.SelectedItem.Key, 1) = "P" Then
            '1.拼音
            For lngCount = 1 To msf.Rows - 1
                If msf.MergeRow(lngCount) Then
                    If msf.TextMatrix(lngCount, 0) = strCode Then
                        For lngLoop = lngCount + 1 To msf.Rows - 1
                            For lngCol = 0 To msf.Cols - 1
                                If msf.TextMatrix(lngLoop, lngCol) = strName Then
                                    msf.Row = lngLoop
                                    msf.Col = lngCol
                                    msf.ShowCell msf.Row, msf.Col
                                    msf.SetFocus
                                    Exit For
                                End If
                            Next
                        Next
                        Exit For
                    End If
                End If
            Next
        Else
            '2.五笔
            For lngLoop = 0 To msf.Rows - 1
                If msf.TextMatrix(lngLoop, 0) = strName Then
                    For lngCol = 1 To msf.Cols - 1
                        If msf.TextMatrix(lngLoop, lngCol) = strCode Then
                            msf.Row = lngLoop
                            msf.Col = lngCol
                            msf.ShowCell msf.Row, msf.Col
                            msf.SetFocus
                            Exit For
                        End If
                    Next
                    Exit For
                End If
            Next
        End If
    Else
        '基本词
        For lngLoop = 1 To msf.Rows - 1
            If msf.TextMatrix(lngLoop, 0) = strName And msf.TextMatrix(lngLoop, 1) = strCode Then
                msf.Row = lngLoop
                msf.ShowCell msf.Row, msf.Col
                msf.SetFocus
                Exit For
            End If
        Next
    End If
    
End Sub

Private Sub EditData(ByVal bytMode As Byte)
    Dim lngLoop As Long
    Dim strWord As String
    Dim strCode As String
    
    Select Case bytMode
    Case 1  '新增
        
        If Left(tvw.SelectedItem.Key, 1) = "P" Then
            Call frmInputBasicEdit.ShowEdit(Me, True, "", "", 1, Val(Left(cbo.Text, 1)))
        Else
            Call frmInputBasicEdit.ShowEdit(Me, True, "", "", 2, Val(Left(cbo.Text, 1)))
        End If
        
    Case 2, 3, 4 '复制新增,修改,删除
        If Trim(msf.TextMatrix(msf.Row, 0)) = "" Then Exit Sub
        
        If Val(Left(cbo.Text, 1)) = 1 Then
            
            '基本字
            If Left(tvw.SelectedItem.Key, 1) = "P" Then
                
                If msf.MergeRow(msf.Row) Then Exit Sub
                
                strWord = msf.TextMatrix(msf.Row, msf.Col)
                For lngLoop = msf.Row - 1 To 1 Step -1
                    If msf.MergeRow(lngLoop) Then
                        strCode = msf.TextMatrix(lngLoop, 0)
                        Exit For
                    End If
                Next
            Else
                strWord = msf.TextMatrix(msf.Row, 0)
                strCode = msf.TextMatrix(msf.Row, msf.Col)
            End If
        Else
            '基本词
            strWord = msf.TextMatrix(msf.Row, 0)
            strCode = msf.TextMatrix(msf.Row, 1)
        End If
        
        If bytMode = 4 Then
            If MsgBox("是否真的要删除当前的编码？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
            
            On Error GoTo ErrHand
            
            gcnOracle.BeginTrans
            If Left(tvw.SelectedItem.Key, 1) = "P" Then
                gcnOracle.Execute "delete from zlWordBasic WHERE 字词='" & strWord & "' AND 输入码='" & strCode & "' and 输入法=1"
            Else
                gcnOracle.Execute "delete from zlWordBasic WHERE 字词='" & strWord & "' AND 输入码='" & strCode & "' and 输入法=2"
            End If
            
            gcnOracle.CommitTrans
            
            If Not (tvw.SelectedItem Is Nothing) Then
                mstrLastKey = ""
                Call tvw_NodeClick(tvw.SelectedItem)
            End If
            
        Else
            If Left(tvw.SelectedItem.Key, 1) = "P" Then
                Call frmInputBasicEdit.ShowEdit(Me, IIf(bytMode = 2, True, False), strWord, strCode, 1, Val(Left(cbo.Text, 1)))
            Else
                Call frmInputBasicEdit.ShowEdit(Me, IIf(bytMode = 2, True, False), strWord, strCode, 2, Val(Left(cbo.Text, 1)))
            End If
        End If
    End Select
    
    Exit Sub
    
ErrHand:
    gcnOracle.RollbackTrans
    MsgBox "删除当前编码失败！" & vbNewLine & Err.Description, vbInformation, gstrSysName
End Sub

Public Sub LocationItem(ByVal strUpKey As String, ByVal strName As String, ByVal strCode As String)
    '定位到指定的资料目录上
    
    On Error Resume Next
    
    tvw.Nodes(strUpKey).Selected = True
    tvw.Nodes(strUpKey).EnsureVisible
    
    If Not (tvw.SelectedItem Is Nothing) Then
        
        Call tvw_NodeClick(tvw.SelectedItem)
        
        Call LocationPostion(strName, strCode)
        
    End If
    
End Sub

Private Sub AdjustMenuEnabled()
    
    mnuFilePrint.Enabled = True
    mnuFilePrintView.Enabled = True
    mnuFileOutExcel.Enabled = True
    
    mnuEditAdd.Enabled = True
    mnuEditAddCopy.Enabled = True
    mnuEditModify.Enabled = True
    mnuEditDelete.Enabled = True
    
    If Trim(msf.TextMatrix(msf.Rows - 1, 0)) = "" And msf.Rows = 2 Then
        mnuFilePrint.Enabled = False
        mnuFilePrintView.Enabled = False
        mnuFileOutExcel.Enabled = False
        mnuEditAddCopy.Enabled = False
        mnuEditModify.Enabled = False
        mnuEditDelete.Enabled = False
    Else
        If Val(msf.Cell(flexcpData, msf.Row, msf.Col)) = 1 Or Trim(msf.TextMatrix(msf.Row, msf.Col)) = "" Then
            mnuEditAddCopy.Enabled = False
            mnuEditModify.Enabled = False
            mnuEditDelete.Enabled = False
        End If
    End If
    
    If Val(Left(cbo.Text, 1)) = 1 Then
        mnuEditAdd.Enabled = False
        mnuEditAddCopy.Enabled = False
        mnuEditModify.Enabled = False
        mnuEditDelete.Enabled = False
    End If
    
    tbrThis.Buttons("预览").Enabled = mnuFilePrintView.Enabled
    tbrThis.Buttons("打印").Enabled = mnuFilePrint.Enabled
    tbrThis.Buttons("增加").Enabled = mnuEditAdd.Enabled
    tbrThis.Buttons("修改").Enabled = mnuEditModify.Enabled
    tbrThis.Buttons("删除").Enabled = mnuEditDelete.Enabled
    
End Sub

Public Sub RefreshData(Optional ByVal strName As String, Optional ByVal strCode As String)

    If tvw.SelectedItem Is Nothing Then Exit Sub
    
    mstrLastKey = ""
    Call tvw_NodeClick(tvw.SelectedItem)
                
    Call LocationPostion(strName, strCode)
    
End Sub
    
Private Sub PrintData(ByVal bytMode As Byte)
    '功能： 打印数据
    '参数： bytMode                         打印方式（1-打印；2-预览；3-输出到Excel）
    
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    If tvw.SelectedItem Is Nothing Then Exit Sub
    
    If msf.TextMatrix(1, 0) = "" Then Exit Sub
    
    If Val(Left(cbo.Text, 1)) = 1 Then
        objPrint.Title = "基本字信息"
    Else
        objPrint.Title = "基本词信息"
    End If
    
    Set objRow = New zlTabAppRow
    
    If Left(tvw.SelectedItem.Key, 1) = "P" Then
        objRow.Add "输入法：拼音   分类：" & tvw.SelectedItem.Text
    Else
        objRow.Add "输入法：五笔   分类：" & tvw.SelectedItem.Text
    End If
    
    objRow.Add ""
    objRow.Add ""
    
    objPrint.UnderAppRows.Add objRow
    
    Set objPrint.Body = msf
        
    If bytMode = 1 Then
        Select Case zlPrintAsk(objPrint)
        Case 1
            zlPrintOrView1Grd objPrint, 1
        Case 2
            zlPrintOrView1Grd objPrint, 2
        Case 3
            zlPrintOrView1Grd objPrint, 3
        End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub AutoRowHeight(ByRef obj As Object)
    '-----------------------------------------------------------------------------------------
    '功能:设置自动行高
    '参数:
    '-----------------------------------------------------------------------------------------
    Call obj.AutoSize(0, obj.Cols - 1)
End Sub

Private Sub cbo_Click()
    Dim objNode As Node
    Dim strTmp As String
    Dim lngLoop As Long
    Dim varArray As Variant
    
    If mlngCboIndex = cbo.ListIndex Then Exit Sub
    mlngCboIndex = cbo.ListIndex
    
    If cbo.ListIndex = 1 Then
        For mlngLoop = tvw.Nodes.Count To 1 Step -1
            
            '移除
            If Left(tvw.Nodes(mlngLoop).Key, 1) = "P" And tvw.Nodes(mlngLoop).Key <> "P0" And Len(tvw.Nodes(mlngLoop).Key) > 2 Then
                tvw.Nodes.Remove mlngLoop
            End If
            
        Next
    Else
        
        For mlngLoop = 1 To 26
            
            strTmp = LCase(Chr(64 + mlngLoop))
            
            If strTmp <> "i" And strTmp <> "u" And strTmp <> "v" Then
                
                On Error Resume Next
                
                Set objNode = tvw.Nodes("P" & strTmp)
                varArray = Split(mstrArray(mlngLoop), ";")
            
                For lngLoop = 0 To UBound(varArray)
                    tvw.Nodes.Add objNode.Key, tvwChild, "P" & strTmp & varArray(lngLoop), varArray(lngLoop), 2, 2
                Next
                
                On Error GoTo 0
                
            End If
        Next
        
    End If
    
    If tvw.SelectedItem Is Nothing Then Exit Sub
    
    mstrLastKey = ""
    Call tvw_NodeClick(tvw.SelectedItem)
End Sub

Private Sub Form_Activate()
    Dim strTmp As String
    Dim lngLoop As Long
    Dim varArray As Variant
    Dim objNode As Node
    
    If mblnStartUp = False Then Exit Sub
    mblnStartUp = False
    
    cbo.AddItem "1 - 基本字"
    cbo.AddItem "2 - 基本词"
    cbo.ListIndex = 0
    
    mnuHelpWeb.Caption = "&Web上的" & gstrWebSustainer
    mnuHelpWebHome.Caption = gstrWebSustainer & "主页"
    Call ApplyOEM(stbThis)
    Call ApplyOEM_Picture(Me, "Icon")
    
    tvw.Nodes.Add , , "P0", "拼音检字表", 1, 1
    
    Set objNode = tvw.Nodes.Add(, , "W0", "五笔检字表", 1, 1)
    
    For mlngLoop = 1 To 26
        strTmp = LCase(Chr(64 + mlngLoop))
        
        If strTmp <> "i" And strTmp <> "u" And strTmp <> "v" Then
            Set objNode = tvw.Nodes.Add("P0", tvwChild, "P" & strTmp, strTmp, 2, 2)
            
'            varArray = Split(mstrArray(mlngLoop), ";")
'
'            For lngLoop = 0 To UBound(varArray)
'                tvw.Nodes.Add objNode.Key, tvwChild, "P" & strTmp & varArray(lngLoop), varArray(lngLoop), 2, 2
'            Next
            
        End If
        
        Set objNode = tvw.Nodes.Add("W0", tvwChild, "W" & strTmp, strTmp, 2, 2)
    Next
    
    mlngCboIndex = -1
    Call cbo_Click
    
    tvw.Nodes("Pa").Selected = True
    mstrLastKey = ""
    Call tvw_NodeClick(tvw.SelectedItem)
    
End Sub

Private Sub Form_Load()
    mblnStartUp = True
    
    mstrArray(1) = "a;ai;an;ang;ao"
    mstrArray(2) = "ba;bai;ban;bang;bao;bei;ben;beng;bi;bian;biao;bie;bin;bing;bo;bu"
    mstrArray(3) = "ca;cai;can;cang;cao;ce;cen;ceng;cha;chai;chan;chang;chao;che;chen;cheng;chi;chong;chou;chu;chua;chuai;chuan;chuang;chui;chun;chuo;ci;cong;cou;cu;cuan;cui;cun;cuo"
    mstrArray(4) = "da;dai;dang;dao;de;dei;den;deng;di;dia;dian;diao;die;ding;diu;dong;dou;du;duan;dui;dun;duo"
    mstrArray(5) = "e;ei;en;eng;er"
    mstrArray(6) = "fa;fan;fang;fei;fen;feng;fo;fou;fu"
    mstrArray(7) = "ga;gai;gan;gang;gao;ge;gei;gen;geng;gong;gou;gu;gua;guan;guang;gui;gun;guo"
    mstrArray(8) = "ha;hai;han;hang;hao;he;hei;hen;heng;hng;hong;hou;hu;hua;huai;huan;huang;hui;hun;huo"
    mstrArray(10) = "ji;jia;jian;jiang;jiao;jie;jin;jing;jiong;jiu;ju;juan;jue;jun"
    mstrArray(11) = "ka;kai;kan;kang;kao;ke;kei;ken;keng;kong;kou;ku;kua;kuai;kuan;kuang;kui;kun;kuo"
    mstrArray(12) = "la;lai;lan;lang;lao;le;lei;leng;li;lia;lian;liang;liao;lie;lin;ling;liu;lo;long;lou;lu;luan;lun;luo;lv;lve"
    mstrArray(13) = "m;ma;mai;man;mang;mao;me;mei;men;meng;mi;mian;miao;mie;min;ming;miu;mo;mou;mu"
    mstrArray(14) = "ngn;na;nai;nan;nang;nao;ne;nei;nen;neng;ng;ni;nian;niang;niao;nie;nin;ning;niu;nong;nou;nu;nuan;nun;nuo;nv;nve"
    mstrArray(15) = "o;ou"
    mstrArray(16) = "pa;pai;pan;pang;pao;pei;pen;peng;pi;pian;piao;pie;pin;ping;po;pou;pu"
    mstrArray(17) = "qi;qia;qian;qiang;qianwa;qiao;qie;qin;qing;qiong;qiu;qu;quan;que;qun"
    mstrArray(18) = "ran;rang;rao;re;ren;reng;ri;rong;rou;ru;ruan;rui;run;ruo"
    mstrArray(19) = "sa;sai;san;sang;sao;se;sen;seng;sha;shai;shan;shang;shao;she;shei;shen;sheng;shi;shou;shu;shua;shuai;shuan;shuang;shui;shun;shuo;si;song;sou;su;suan;sui;sun;suo"
    mstrArray(20) = "ta;tai;tan;tang;tao;te;teng;ti;tian;tiao;tie;ting;tong;tou;tu;tuan;tui;tun;tuo"
    mstrArray(23) = "wa;wai;wan;wang;wei;wen;weng;wo;wu"
    mstrArray(24) = "xi;xia;xian;xiang;xiao;xie;xin;xing;xiong;xu;xuan;xue;xun"
    mstrArray(25) = "ya;yan;yang;yao;ye;yi;yin;ying;yingli;yo;yong;you;yu;yuan;yue;yun"
    mstrArray(26) = "za;zai;zan;zang;zao;ze;zei;zen;zeng;zha;zhai;zhan;zhang;zhao;zhe;zhei;zhen;zheng;zhi;zhong;zhou;zhu;zhua;zhuai;zhuan;zhuang;zhui;zhun;zhuo;zi;zong;zou;zu;zuan;zui;zun;zuo"
    
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    With tvw
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = imgY.Left
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    
    With imgY
        .Top = tvw.Top
        .Height = tvw.Height
    End With
    
    With msf
        .Left = imgY.Left + imgY.Width
        .Top = tvw.Top
        .Width = Me.ScaleWidth - .Left
        .Height = tvw.Height
    End With
    
End Sub

Private Sub imgY_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    imgY.Left = imgY.Left + X
    
    If imgY.Left < 1500 Then imgY.Left = 1500
    If Me.Width - imgY.Left - imgY.Width < 1000 Then imgY.Left = Me.Width - imgY.Width - 1000

    Form_Resize
End Sub

Private Sub mnuEditAdd_Click()
    Call EditData(1)
End Sub

Private Sub mnuEditAddCopy_Click()
    Call EditData(2)
End Sub

Private Sub mnuEditDelete_Click()
    Call EditData(4)
End Sub

Private Sub mnuEditModify_Click()
    Call EditData(3)
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileOutExcel_Click()
    Call PrintData(3)
End Sub

Private Sub mnuFilePrint_Click()
    Call PrintData(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuFilePrintView_Click()
    Call PrintData(2)
End Sub

Private Sub mnuHelpAbout_Click()
    frmAbout.ShowAbout
End Sub

Private Sub mnuHelpTopic_Click()
    ShowHelp Me.hwnd, "zl9svrtools\" & Me.name
End Sub

Private Sub mnuHelpWebHome_Click()
    ShellExecute hwnd, "open", "http://" & gstrWebURL, "", "", 1
End Sub

Private Sub mnuHelpWebMail_Click()
    ShellExecute hwnd, "open", "mailto:" & gstrWebEmail, "", "", 1
End Sub

Private Sub mnuViewFind_Click()
    frmInputFind.ShowFind Me, Val(Left(cbo.Text, 1))
End Sub

Private Sub mnuViewRefresh_Click()
    Dim strSvrName As String
    Dim strSvrCode As String
    
    If tvw.SelectedItem Is Nothing Then Exit Sub
    
    Call SavePostion(Val(Left(cbo.Text, 1)), strSvrName, strSvrCode)
    
    mstrLastKey = ""
    Call tvw_NodeClick(tvw.SelectedItem)
    
    Call LocationPostion(strSvrName, strSvrCode)
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

Private Sub msf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AutoRowHeight(msf)
End Sub

Private Sub msf_DblClick()
   
    If Trim(msf.TextMatrix(msf.Row, 0)) <> "" Then
        If mnuEdit.Visible And mnuEditModify.Visible And mnuEditModify.Enabled Then
            Call mnuEditModify_Click
        End If
    End If

End Sub

Private Sub msf_EnterCell()
    If msf.TextMatrix(msf.Row, msf.Col) <> "" Then
        msf.FocusRect = flexFocusSolid
        msf.HighLight = flexHighlightAlways
    Else
        msf.FocusRect = flexFocusNone
        msf.HighLight = flexHighlightNever
    End If
    
    Call AdjustMenuEnabled
End Sub

Private Sub msf_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call msf_DblClick
End Sub

Private Sub msf_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And mnuEdit.Visible Then Me.PopupMenu mnuEdit
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "预览"
        Call mnuFilePrintView_Click
    Case "打印"
        Call mnuFilePrint_Click
    Case "增加"
        Call mnuEditAdd_Click
    Case "修改"
        Call mnuEditModify_Click
    Case "删除"
        Call mnuEditDelete_Click
    Case "查找"
        Call mnuViewFind_Click
    Case "帮助"
        Call mnuHelpTopic_Click
    Case "退出"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim rs As New ADODB.Recordset
On Error GoTo errHandle
    
    If mstrLastKey = Node.Key Then Exit Sub
    mstrLastKey = Node.Key
    
    If Val(Left(cbo.Text, 1)) = 1 Then
        msf.Rows = 1
        For mlngLoop = 0 To msf.Cols - 1
            msf.TextMatrix(0, mlngLoop) = ""
        Next
    Else
        msf.Rows = 2
        For mlngLoop = 0 To msf.Cols - 1
            msf.TextMatrix(1, mlngLoop) = ""
        Next
    End If
    
    stbThis.Panels(2).Text = ""
    
    If Node.Key = "W0" Or Node.Key = "P0" Then
        Call AdjustMenuEnabled
        Exit Sub
    End If
    
    If Left(Node.Key, 1) = "W" Then
        gstrSQL = "SELECT 字词,输入码 FROM zlWordBasic WHERE 输入法=2 and 是否字=" & Left(cbo.Text, 1) & " and 输入码 Like '" & LCase(Node.Text) & "%' ORDER BY 字词"
    Else
        gstrSQL = "SELECT 字词,输入码 FROM zlWordBasic WHERE 输入法=1 and 是否字=" & Left(cbo.Text, 1) & " and 输入码 Like '" & LCase(Node.Text) & "%' ORDER BY 输入码"
    End If
    
    rs.Open gstrSQL, gcnOracle
    
    If Left(cbo.Text, 1) = 1 Then
        If Left(Node.Key, 1) = "W" Then
            Call ShowLayer_WB(rs)
        Else
            Call ShowLayer_PY(rs)
        End If
    Else
        Call ShowList(rs)
        Call AutoRowHeight(msf)
    End If
    
    Call AdjustMenuEnabled
    
    If Left(Node.Key, 1) = "P" Then
        stbThis.Panels(2).Text = "拼音分类'" & tvw.SelectedItem.Text & "'下共有 " & rs.RecordCount & " 条基本字词。"
    Else
        stbThis.Panels(2).Text = "五笔分类'" & tvw.SelectedItem.Text & "'下共有 " & rs.RecordCount & " 条基本字词。"
    End If

    Exit Sub
errHandle:
    MsgBox Err.Description, vbCritical, Me.Caption
End Sub

Private Sub ShowLayer_PY(ByVal rs As ADODB.Recordset)
    Dim varArray As Variant
    
    Dim lngLoop As Long
    Dim lngCols As Long
    
    msf.Cols = 30
    msf.Rows = 1
    msf.FixedRows = 0
    
    For lngLoop = 0 To msf.Cols - 1
        msf.ColWidth(lngLoop) = 255
    Next
    msf.GridLines = flexGridNone
    msf.SheetBorder = &H80000009
    msf.SelectionMode = flexSelectionFree
    msf.MergeCells = flexMergeFree
    For lngCols = 0 To 3
        msf.TextMatrix(0, lngCols) = ""
        msf.Cell(flexcpData, 0, lngCols) = 0
    Next
    
    
    varArray = Split(mstrArray(Asc(tvw.SelectedItem.Text) - 96), ";")
    
    
    For lngLoop = 0 To UBound(varArray)
        rs.Filter = ""
        
        rs.Filter = "输入码='" & varArray(lngLoop) & "'"
        If rs.RecordCount > 0 Then
            rs.MoveFirst
                        
            If msf.Rows > 1 Then msf.Rows = msf.Rows + 1
            msf.Rows = msf.Rows + 1
                        
            msf.MergeRow(msf.Rows - 1) = True
            For lngCols = 0 To 3
                msf.TextMatrix(msf.Rows - 1, lngCols) = varArray(lngLoop)
                msf.Cell(flexcpData, msf.Rows - 1, lngCols) = 1
            Next
            
            
            lngCols = 0
            msf.Rows = msf.Rows + 1
            Do While Not rs.EOF
                If lngCols > msf.Cols - 1 Then
                    msf.Rows = msf.Rows + 1
                    lngCols = 0
                End If
                
                msf.TextMatrix(msf.Rows - 1, lngCols) = rs("字词").Value
                lngCols = lngCols + 1
                
                rs.MoveNext
            Loop
            
        End If
    Next
End Sub

Private Sub ShowLayer_WB(ByVal rs As ADODB.Recordset)
    Dim varArray As Variant
    Dim strWord As String
    Dim lngLoop As Long
    Dim lngCols As Long
    
    msf.Cols = 15
    msf.Rows = 1
    msf.FixedRows = 0
    msf.FocusRect = flexFocusSolid
    
    For lngLoop = 0 To msf.Cols - 1
        msf.ColWidth(lngLoop) = 550
    Next
    msf.GridLines = flexGridNone
    msf.SheetBorder = &H80000009
    msf.SelectionMode = flexSelectionFree

    lngCols = 0
    Do While Not rs.EOF
    
        If strWord <> rs("字词").Value Then
           strWord = rs("字词").Value
           If msf.TextMatrix(msf.Rows - 1, 0) <> "" Then msf.Rows = msf.Rows + 1
           msf.TextMatrix(msf.Rows - 1, 0) = strWord
           msf.Cell(flexcpData, msf.Rows - 1, 0) = 1
           lngCols = 0
        End If
        
        If lngCols + 1 < 14 Then
            lngCols = lngCols + 1
            msf.TextMatrix(msf.Rows - 1, lngCols) = rs("输入码").Value
        End If
        
        rs.MoveNext
    Loop
End Sub

Private Sub ShowList(ByVal rs As ADODB.Recordset)
    Dim lngCols As Long
    
    msf.Cols = 2
    msf.Rows = 2
    msf.FixedRows = 1
    msf.FocusRect = flexFocusHeavy
    msf.GridLines = flexGridFlat
    msf.SheetBorder = &H8000000C
    msf.SelectionMode = flexSelectionByRow
    
    msf.TextMatrix(0, 0) = "基本词"
    msf.TextMatrix(0, 1) = "输入码"
    
    For lngCols = 0 To msf.Cols - 1
        msf.TextMatrix(1, lngCols) = ""
        msf.Cell(flexcpData, 1, lngCols) = 1
    Next
    
    If rs.BOF = False Then
        Set msf.DataSource = rs
    End If
    
    msf.ColWidth(0) = 1500
    msf.ColWidth(1) = 3000
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

