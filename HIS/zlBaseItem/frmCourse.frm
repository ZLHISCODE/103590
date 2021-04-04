VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCourse 
   Caption         =   "期间划分调整"
   ClientHeight    =   6045
   ClientLeft      =   255
   ClientTop       =   645
   ClientWidth     =   5625
   Icon            =   "frmCourse.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6045
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgStard 
      Left            =   2985
      Top             =   555
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
            Picture         =   "frmCourse.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCourse.frx":0524
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCourse.frx":073E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCourse.frx":0958
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCourse.frx":0B72
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCourse.frx":0D8C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgHot 
      Left            =   3780
      Top             =   570
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
            Picture         =   "frmCourse.frx":0FA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCourse.frx":11C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCourse.frx":13DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCourse.frx":15F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCourse.frx":180E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCourse.frx":1A28
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrTop 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   5625
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinWidth1       =   495
      MinHeight1      =   720
      Width1          =   4305
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   30
         TabIndex        =   3
         Top             =   30
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgStard"
         HotImageList    =   "imgHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "PrintView"
               Object.ToolTipText     =   "预览期间表"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Object.ToolTipText     =   "打印期间表"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "增加"
               Key             =   "Add"
               Object.ToolTipText     =   "增加一期间"
               Object.Tag             =   "增加"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "Delete"
               Object.ToolTipText     =   "删除最大期间"
               Object.Tag             =   "删除"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Object.ToolTipText     =   "帮助主题"
               Object.Tag             =   "帮助"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
   End
   Begin MSComCtl2.DTPicker dtp日期 
      Height          =   330
      Left            =   2010
      TabIndex        =   1
      Top             =   2857
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   582
      _Version        =   393216
      CustomFormat    =   "yyyy年MM月dd日"
      Format          =   98500611
      CurrentDate     =   36179
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hgdPeriod 
      Height          =   3885
      Left            =   120
      TabIndex        =   0
      Top             =   1020
      Width           =   4560
      _ExtentX        =   8043
      _ExtentY        =   6853
      _Version        =   393216
      FixedCols       =   0
      BackColorSel    =   8421376
      AllowBigSelection=   0   'False
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   5685
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   635
      SimpleText      =   $"frmCourse.frx":1C42
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCourse.frx":1C89
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4842
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
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuPrintView 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFileSpt3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "增加一个期间(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除最后期间(&D)"
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "报表(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuReportItem 
         Caption         =   "-"
         Index           =   0
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
         Begin VB.Menu mnuViewToolspilt1 
            Caption         =   "-"
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
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web上的中联"
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
      Begin VB.Menu mnuHelpSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmCourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private rsPeriod As New ADODB.Recordset
Private mblnLoad As Boolean
Private mlngMode As Long
Private mstrPrivs As String                              '权限串

Private Function InitTable() As Boolean
    Err = 0
    Dim intTop As Long
    
    On Error GoTo ErrHand
    With Me.hgdPeriod
        .redraw = False
        .Clear
        .Cols = 3
        .ColWidth(0) = 1600
        .ColWidth(1) = Me.dtp日期.Width + 30
        .ColWidth(2) = Me.dtp日期.Width + 30
        
        .ColAlignment(0) = 4
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        
        .ColAlignmentFixed(0) = 4
        .ColAlignmentFixed(1) = 4
        .ColAlignmentFixed(2) = 4
        
        .TextMatrix(0, 0) = "期间"
        .TextMatrix(0, 1) = "开始日期"
        .TextMatrix(0, 2) = "终止日期"
        
        gstrSQL = "select 期间,开始日期,终止日期," & _
                   " sign(to_number(to_char(sysdate,'YYYYMMDD'))-to_number(to_char(终止日期,'YYYYMMDD'))) as 过去," & _
                   " sign(to_number(to_char(开始日期,'YYYYMMDD'))-to_number(to_char(sysdate,'YYYYMMDD'))) as 未来" & _
                   " from 期间表 order by 期间"
        Set rsPeriod = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

        .RowHeight(0) = Me.dtp日期.Height
        If rsPeriod.RecordCount = 0 Then
            .Rows = 2
            Exit Function
        Else
            .Rows = rsPeriod.RecordCount + 1
        End If
        rsPeriod.MoveFirst
        Do While Not rsPeriod.EOF
            .RowHeight(rsPeriod.AbsolutePosition) = Me.dtp日期.Height
            If rsPeriod.Fields("未来").Value < 1 Then '不是未来
                .Row = rsPeriod.AbsolutePosition
                .Col = 0
                .CellBackColor = IIF(rsPeriod.Fields("过去").Value = 1, &HE0E0E0, &HC0E0FF)
                .Col = 1
                .CellBackColor = IIF(rsPeriod.Fields("过去").Value = 1, &HE0E0E0, &HC0E0FF)
                .Col = 2
                .CellBackColor = IIF(rsPeriod.Fields("过去").Value = 1, &HE0E0E0, &HC0E0FF)
                If rsPeriod.Fields("过去").Value <> 1 Then
                    '表示当前期间
                    intTop = .Row
                End If
            End If
                
            .RowData(rsPeriod.AbsolutePosition) = rsPeriod.Fields("未来").Value
            .TextMatrix(rsPeriod.AbsolutePosition, 0) = Left(rsPeriod.Fields("期间").Value, 4) & "年" & Right(rsPeriod.Fields("期间").Value, 2) & "月"
            .TextMatrix(rsPeriod.AbsolutePosition, 1) = Format(rsPeriod.Fields("开始日期").Value, "yyyy年MM月dd日")
            .TextMatrix(rsPeriod.AbsolutePosition, 2) = Format(rsPeriod.Fields("终止日期").Value, "yyyy年MM月dd日")
            rsPeriod.MoveNext
        Loop
        .redraw = True
        .TopRow = IIF(intTop = 0, 1, intTop)
    End With
    InitTable = True
    
    If intTop = 0 Then
        MsgBox "没有发现当前期间，请增加！", vbExclamation, gstrSysName
    End If
    
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub dtp日期_Change()
    Dim str期间 As String
    
    On Error GoTo ErrHand
    With Me.hgdPeriod
        gcnOracle.BeginTrans
        
        If Format(Me.dtp日期.Value, "YYYY-MM-DD") < Format(Me.dtp日期.MinDate, "YYYY-MM-DD") Then Me.dtp日期.Value = Format(Me.dtp日期.MinDate, "YYYY-MM-DD")
        .TextMatrix(.Row, 2) = Format(Me.dtp日期.Value, "yyyy年MM月dd日")
        gstrSQL = "zl_期间表_update('" & Left(.TextMatrix(.Row, 0), 4) & Mid(.TextMatrix(.Row, 0), 6, 2) & "',null,to_date('" & Format(Me.dtp日期.Value, "YYYY-MM-DD") & "','YYYY-MM-DD'))"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        
        If .Row <> .Rows - 1 Then
            .TextMatrix(.Row + 1, 1) = Format(Me.dtp日期.Value + 1, "yyyy年MM月dd日")
            If Mid(.TextMatrix(.Row, 0), 6, 2) = "12" Then
                str期间 = CStr(Val(Left(.TextMatrix(.Row, 0), 4)) + 1) & "01"
            ElseIf Val(Mid(.TextMatrix(.Row, 0), 6, 2)) >= 9 Then
                str期间 = Left(.TextMatrix(.Row, 0), 4) & CStr(Val(Mid(.TextMatrix(.Row, 0), 6, 2)) + 1)
            Else
                str期间 = Left(.TextMatrix(.Row, 0), 4) & "0" & CStr(Val(Mid(.TextMatrix(.Row, 0), 6, 2)) + 1)
            End If
            gstrSQL = "zl_期间表_update('" & str期间 & "',to_date('" & Format(Me.dtp日期.Value + 1, "YYYY-MM-DD") & "','YYYY-MM-DD'),null)"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
        gcnOracle.CommitTrans
    End With
    Exit Sub

ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dtp日期_LostFocus()
    Me.dtp日期.Visible = False
End Sub

Private Sub Form_Activate()
    If mblnLoad = True Then
        If InitTable = False Then Exit Sub
    End If
    mblnLoad = False
End Sub

Private Sub Form_Load()
    mlngMode = glngModul
    mstrPrivs = gstrPrivs

    RestoreWinState Me, App.ProductName
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    mblnLoad = True
    With rsPeriod
        If .State = adStateOpen Then .Close
    End With
    
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    SizeControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me, App.ProductName
End Sub

Private Sub hgdPeriod_DblClick()
    If Me.hgdPeriod.Col = 2 Then
        hgdPeriod_EnterCell
    End If
End Sub

Private Sub hgdPeriod_EnterCell()
    Dim dtStart As Date
    If Me.hgdPeriod.RowData(Me.hgdPeriod.Row) > 0 And Me.hgdPeriod.Col = 2 Then
        dtStart = Left(Me.hgdPeriod.TextMatrix(Me.hgdPeriod.Row, 0), 4) & "-" & Mid(Me.hgdPeriod.TextMatrix(Me.hgdPeriod.Row, 0), 6, 2) & "-1"
        Me.dtp日期.MinDate = 0
        Me.dtp日期.MaxDate = CDate("9999-12-31")
        Me.dtp日期.MinDate = dtStart + 19
        Me.dtp日期.MaxDate = dtStart + 40
        Me.dtp日期.Value = Me.hgdPeriod.TextMatrix(Me.hgdPeriod.Row, 2)
        Me.dtp日期.Move Me.hgdPeriod.Left + Me.hgdPeriod.CellLeft - 30, Me.hgdPeriod.Top + Me.hgdPeriod.CellTop, Me.hgdPeriod.CellWidth + 45
        Me.dtp日期.Visible = True
    Else
        Me.dtp日期.Move 0, 0
        Me.dtp日期.Visible = False
    End If
End Sub

Private Sub hgdPeriod_GotFocus()
    Me.dtp日期.Visible = False
    Me.dtp日期.Move 0, 0
End Sub

Private Sub hgdPeriod_Scroll()
    Me.dtp日期.Visible = False
    Me.dtp日期.Move 0, 0
End Sub

Private Sub SizeControls()
    Dim intTop As Integer, intButton As Integer
    intTop = IIF(Me.cbrTop.Visible, Me.cbrTop.Height, 0)
    intButton = IIF(Me.stbThis.Visible, Me.stbThis.Height, 0)
    On Error Resume Next
    
    With Me.hgdPeriod
        .Top = intTop
        .Left = Me.ScaleLeft
        .Height = Me.ScaleHeight - intTop - intButton
        .Width = Me.ScaleWidth
    End With
End Sub

Private Sub mnuEditAdd_Click()
    Dim strMonth As String
    
    On Error GoTo ErrHandle
    With Me.hgdPeriod
        .Row = .Rows - 1
        If Mid(.TextMatrix(.Row, 0), 6, 2) = "12" Then
            strMonth = CStr(Val(Left(.TextMatrix(.Row, 0), 4)) + 1) & "01"
        ElseIf Val(Mid(.TextMatrix(.Row, 0), 6, 2)) >= 9 Then
            strMonth = Left(.TextMatrix(.Row, 0), 4) & CStr(Val(Mid(.TextMatrix(.Row, 0), 6, 2)) + 1)
        Else
            strMonth = Left(.TextMatrix(.Row, 0), 4) & "0" & CStr(Val(Mid(.TextMatrix(.Row, 0), 6, 2)) + 1)
        End If
        .Rows = .Rows + 1
        .RowHeight(.Rows - 1) = Me.dtp日期.Height
        .RowData(.Rows - 1) = 1
        .TextMatrix(.Rows - 1, 0) = Left(strMonth, 4) & "年" & Right(strMonth, 2) & "月"
        .TextMatrix(.Rows - 1, 1) = Format(CDate(.TextMatrix(.Rows - 2, 2)) + 1, "yyyy年MM月dd日")
        
        .TextMatrix(.Rows - 1, 2) = Format(DateAdd("m", 1, CDate(.TextMatrix(.Rows - 2, 2))), "yyyy年MM月dd日")
        .TextMatrix(.Rows - 1, 2) = Format(DateSerial(Year(CDate(.TextMatrix(.Rows - 1, 2))), Month(CDate(.TextMatrix(.Rows - 1, 2))) + 1, 0), "yyyy年MM月dd日")
        
        gstrSQL = "zl_期间表_insert('" & strMonth & "', to_date('" & _
                  Format(CDate(.TextMatrix(.Rows - 1, 1)), "yyyy-mm-dd") & "','YYYY-MM-DD'), to_date('" & _
                  Format(CDate(.TextMatrix(.Rows - 1, 2)), "yyyy-mm-dd") & "','YYYY-MM-DD'))"
        
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
'        rsPeriod.Requery
        .TopRow = .Rows - 1
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditDelete_Click()
    With Me.hgdPeriod
        If .RowData(.Rows - 1) <= 0 Then
            MsgBox "不能删除当前期间(只能删除未来的期间)。", vbExclamation, gstrSysName
            Exit Sub
        End If
        If MsgBox("真的要删除期间 <" & .TextMatrix(.Rows - 1, 0) & "> 吗？", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        gstrSQL = "zl_期间表_delete('" & Left(.TextMatrix(.Rows - 1, 0), 4) & Mid(.TextMatrix(.Rows - 1, 0), 6, 2) & "')"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
'        rsPeriod.Requery
        .Rows = .Rows - 1
        .TopRow = .Rows - 1
    End With
End Sub

Private Sub mnuFileExcel_Click()
    Dim objPrint As New zlPrint1Grd
    objPrint.Title.Text = "期间划分表"
    Set objPrint.Body = Me.hgdPeriod
    zlPrintOrView1Grd objPrint, 3

End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePrint_Click()
    Dim objPrint As New zlPrint1Grd
    objPrint.Title.Text = "期间划分表"
    Set objPrint.Body = Me.hgdPeriod
    Select Case zlPrintAsk(objPrint)
    Case 1
        zlPrintOrView1Grd objPrint, 1
    Case 2
        zlPrintOrView1Grd objPrint, 2
    End Select
End Sub

Private Sub mnuFilePrintset_Click()
    zlPrintSet
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpHelp_Click()
    ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hwnd)
End Sub

Private Sub mnuPrintView_Click()
    Dim objPrint As New zlPrint1Grd
    objPrint.Title.Text = "期间划分表"
    Set objPrint.Body = Me.hgdPeriod
    zlPrintOrView1Grd objPrint, 2
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me)
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = Me.mnuViewStatus.Checked
    SizeControls
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not Me.mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    Me.cbrTop.Visible = Me.mnuViewToolButton.Checked
    SizeControls
End Sub

Private Sub mnuViewToolText_Click()
    Dim i As Integer
    Me.mnuViewToolText.Checked = Not Me.mnuViewToolText.Checked
    If Me.mnuViewToolText.Checked Then
        For i = 1 To Me.tbrThis.Buttons.Count
            Me.tbrThis.Buttons(i).Caption = Me.tbrThis.Buttons(i).Tag
        Next
    Else
        For i = 1 To Me.tbrThis.Buttons.Count
            Me.tbrThis.Buttons(i).Caption = ""
        Next
    End If
    Me.cbrTop.Bands(1).MinHeight = Me.tbrThis.Height
    Me.cbrTop.Refresh
    SizeControls

End Sub


Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
    Case "PrintView"
        mnuPrintView_Click
    Case "Print"
        mnuFilePrint_Click
    Case "Add"
        mnuEditAdd_Click
    Case "Delete"
        mnuEditDelete_Click
    Case "Help"
        mnuHelpHelp_Click
    Case "Exit"
        mnuFileExit_Click
    End Select
    
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool, 2
    End If
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

