VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmTrace 
   AutoRedraw      =   -1  'True
   Caption         =   "Trace"
   ClientHeight    =   5250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   Icon            =   "frmTrace.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5250
   ScaleWidth      =   7005
   WindowState     =   2  'Maximized
   Begin VB.Frame fraFilter 
      Appearance      =   0  'Flat
      BackColor       =   &H00E1FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   60
      TabIndex        =   10
      Top             =   90
      Visible         =   0   'False
      Width           =   6885
      Begin ZLSQLTrace.ccXPButton cmdFilter 
         Height          =   360
         Left            =   5940
         TabIndex        =   3
         Top             =   45
         Width           =   855
         _extentx        =   1508
         _extenty        =   635
         caption         =   "确定(&O)"
         font            =   "frmTrace.frx":038A
      End
      Begin VB.TextBox txtShoot 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   4635
         MaxLength       =   2
         TabIndex        =   2
         Top             =   90
         Width           =   360
      End
      Begin VB.CheckBox chkFull 
         BackColor       =   &H00E1FFFF&
         Caption         =   "仅显示全表扫描和索引全扫描"
         Height          =   195
         Left            =   255
         TabIndex        =   0
         Top             =   135
         Width           =   2640
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "缓存命中率低于     %"
         Height          =   180
         Left            =   3330
         TabIndex        =   1
         Top             =   135
         Width           =   1800
      End
      Begin VB.Line Line1 
         X1              =   90
         X2              =   660
         Y1              =   390
         Y2              =   390
      End
   End
   Begin MSComctlLib.ImageList imgCaption 
      Left            =   2985
      Top             =   3750
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
            Picture         =   "frmTrace.frx":03B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTrace.frx":074C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraUD 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   4035
      MousePointer    =   7  'Size N S
      TabIndex        =   9
      Top             =   2490
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Frame fraLR 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4470
      Left            =   2640
      MousePointer    =   9  'Size W E
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.TextBox txtSQL 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   1770
      IMEMode         =   2  'OFF
      Left            =   2715
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   645
      Visible         =   0   'False
      Width           =   4245
   End
   Begin VSFlex8Ctl.VSFlexGrid vsPlan 
      Height          =   1935
      Left            =   2700
      TabIndex        =   6
      Top             =   2550
      Visible         =   0   'False
      Width           =   4260
      _cx             =   7514
      _cy             =   3413
      Appearance      =   2
      BorderStyle     =   0
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
      BackColorFixed  =   -2147483643
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
      FocusRect       =   0
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   235
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmTrace.frx":0AE6
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
      OutlineBar      =   4
      OutlineCol      =   1
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
   Begin RichTextLib.RichTextBox txtTrace 
      Height          =   1590
      Left            =   390
      TabIndex        =   7
      Top             =   3375
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   2805
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      Appearance      =   0
      TextRTF         =   $"frmTrace.frx":0B2E
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
   Begin VSFlex8Ctl.VSFlexGrid vsTrace 
      Height          =   2520
      Left            =   0
      TabIndex        =   4
      Top             =   630
      Visible         =   0   'False
      Width           =   2625
      _cx             =   4630
      _cy             =   4445
      Appearance      =   2
      BorderStyle     =   0
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
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   235
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmTrace.frx":0BCB
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
End
Attribute VB_Name = "frmTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event UpdateStatus(ByVal strStatus As String)
Private mstrFile As String
Private mstrSort As String

Private WithEvents mfrmFind As frmFind
Attribute mfrmFind.VB_VarHelpID = -1
Private mstrFind As String
Private mblnMachCase As Boolean
Private mblnMultiRows As Boolean

Private mlngMinSize As Long '中型表大小
Private mlngMaxSize As Long
Private mrsBigTbl As ADODB.Recordset    '需要检查的表
Private mrsBigIdx As ADODB.Recordset
Private mrsLowIdx As ADODB.Recordset

Private Const COLOR_FULL = &HF0F0FF
Private Enum Col_Trace
    COL_标题 = 0
    COL_次数 = 1
    COL_CPU时间 = 2
    COL_总时间 = 3
    COL_物理读 = 4
    COL_一致读 = 5
    COL_当前读 = 6
    COL_记录数 = 7
    COL_命中率 = 8
End Enum

Private Type SQLError
    SQL As String
    Err As String
End Type

Public mlngCount As Long    '本窗体暂时未使用

Public Property Get Filtering() As Boolean
    Filtering = fraFilter.Visible = True
End Property

Public Property Get ViewStyle() As CommandBarIDCond
    If txtTrace.Visible Then
        ViewStyle = conMenu_View_Style_Report
    Else
        ViewStyle = conMenu_View_Style_Table
    End If
End Property

Public Sub ShowMe(frmMain As Object, ByVal strFile As String)
    mstrFile = strFile
    mlngCount = 0
    Me.Show
End Sub

Public Sub DoCommand(ByVal DoID As CommandBarIDCond)
'功能：子窗体命令执行接口
    Dim lngRow As Long, i As Long, k As Long
    Dim strL As String, strR As String
    Dim vSel As CHARRANGE
    
    If DoID = conMenu_View_Style Then
        If ViewStyle = conMenu_View_Style_Report Then
            DoID = conMenu_View_Style_Table
        ElseIf ViewStyle = conMenu_View_Style_Table Then
            DoID = conMenu_View_Style_Report
        End If
    End If
    
    If DoID = conMenu_Edit_CompareLeft Then
        gstrLeft = mstrFile
    ElseIf DoID = conMenu_Edit_Compare Then
        '处理短文件路径
        strL = GetShortName(gobjFile.GetParentFolderName(gstrLeft)) & "\" & gobjFile.GetFileName(gstrLeft)
        strR = GetShortName(gobjFile.GetParentFolderName(mstrFile)) & "\" & gobjFile.GetFileName(mstrFile)
        Err.Clear: On Error Resume Next
        Shell gstrCompareExe & " " & strL & " " & strR & " /r /noedit /readonly /fv", vbNormalFocus
        If Err.Number = 0 Then gstrLeft = ""
        Err.Clear
    ElseIf DoID = conMenu_View_Style_Report Then
        Set Me.Icon = imgCaption.ListImages(1).Picture
        fraFilter.Visible = False
        txtTrace.Visible = True: vsTrace.Visible = False
        txtSQL.Visible = False: vsPlan.Visible = False
        fraLR.Visible = False: fraUD.Visible = False
        Call Form_Resize
        
        '根据表格当前内容定位报告文件行
        If vsTrace.Rows <> 0 Then
            lngRow = GetBaseRow(vsTrace.Row)
            lngRow = vsTrace.Cell(flexcpData, lngRow, vsTrace.Cols - 1) - 1
        End If
        k = SendMessage(txtTrace.hwnd, EM_LINEINDEX, lngRow, 0)
        If k <> -1 Then
            'txtTrace.SelStart = Len(txtTrace.Text) '目的是让选中行成为顶行(很慢)
            vSel.cpMin = k: vSel.cpMax = k
            SendMessage txtTrace.hwnd, EM_EXSETSEL, 0, vSel
            
            'SendMessage txtTrace.hWnd, EM_SETSEL, k, k'Selection End不对,要一直到最后。
        End If
    ElseIf DoID = conMenu_View_Style_Table Then
        mfrmFind.Hide
        If vsTrace.Rows = 0 Then Call FileToTable
        
        Set Me.Icon = imgCaption.ListImages(2).Picture
        fraFilter.Visible = True
        txtTrace.Visible = False: vsTrace.Visible = True
        txtSQL.Visible = True: vsPlan.Visible = True
        fraLR.Visible = True: fraUD.Visible = True
        Call Form_Resize
        
        If vsTrace.Rows = 0 Then Exit Sub
        
        '根据报告当前内容定位表格
        lngRow = SendMessage(txtTrace.hwnd, EM_LINEINDEX, -1, 0)
        lngRow = SendMessage(txtTrace.hwnd, EM_LINEFROMCHAR, lngRow, 0) + 1
        With vsTrace
            k = -1
            For i = .FixedRows To .Rows - 1 Step 5
                If .Cell(flexcpData, i, .Cols - 1) <= lngRow Then
                    k = i
                Else
                    Exit For
                End If
            Next
            If k <> -1 Then
                If Not .RowHidden(k) Then
                    .Row = k: .Col = 0
                    .ShowCell .Row + 4, .Col
                    If .Row < .TopRow Then
                        .ShowCell .Row, .Col
                    End If
                End If
            End If
        End With
    ElseIf DoID = conMenu_View_Find Then
        mfrmFind.ShowMe txtTrace.SelText
    ElseIf DoID = conMenu_View_FindNext Then
        If mstrFind = "" Then
            mfrmFind.ShowMe txtTrace.SelText
        Else
            Call SearchText
            txtTrace.SetFocus
        End If
    ElseIf DoID = conMenu_View_Filter Then
        fraFilter.Visible = Not fraFilter.Visible
        Call Form_Resize
        If fraFilter.Visible Then chkFull.SetFocus
    ElseIf DoID = conMenu_View_SQLPrev Then
        With vsTrace
            If .Row = -1 Then
                lngRow = GetBaseRow(.TopRow)
                .Row = lngRow: Call .ShowCell(.Row, .Col)
            Else
                lngRow = GetBaseRow(.Row)
                For i = lngRow - 1 To .FixedRows Step -1
                    If .RowData(i) > 0 And Not .RowHidden(i) Then
                        .Row = i: Call .ShowCell(.Row, .Col): Exit For
                    End If
                Next
            End If
        End With
    ElseIf DoID = conMenu_View_SQLNext Then
        With vsTrace
            If .Row = -1 Then
                lngRow = GetBaseRow(.TopRow)
                .Row = lngRow: Call .ShowCell(.Row, .Col)
            Else
                lngRow = GetBaseRow(.Row)
                For i = lngRow + 1 To .Rows - 1
                    If .RowData(i) > 0 And Not .RowHidden(i) Then
                        .Row = i: Call .ShowCell(.Row + 4, .Col): Exit For
                    End If
                Next
            End If
        End With
    ElseIf DoID = conMenu_View_Close Then
        Unload Me
    End If
End Sub

Public Function GetCommand(ByVal DoID As CommandBarIDCond) As Boolean
'功能：子窗体命令状态接口
    Select Case DoID
    Case conMenu_Edit_CompareLeft
        GetCommand = True
    Case conMenu_View_Style
        GetCommand = True
    Case conMenu_View_Find
        GetCommand = ViewStyle = conMenu_View_Style_Report
    Case conMenu_View_FindNext
        GetCommand = ViewStyle = conMenu_View_Style_Report And mstrFind <> ""
    Case conMenu_View_Filter
        GetCommand = Me.ViewStyle = conMenu_View_Style_Table
    Case conMenu_View_SQLPrev, conMenu_View_SQLNext
        If vsTrace.Visible And vsTrace.Rows > 0 Then
            GetCommand = True
        End If
    End Select
End Function

Private Sub SearchText()
    Static blnStart As Boolean
    Dim k As Long, vFind As FINDTEXT, vSel As CHARRANGE
        
    '自带的查找:
    '功能有时不正确,不向下继续查找,甚至向上查找
    '但自带的Find方式可自动选择定位
    'k = txtTrace.Find(mstrFind, IIf(blnStart, 0, txtTrace.SelStart + txtTrace.SelLength), , IIf(mblnMachCase, rtfMatchCase, 0))
    
    'API查找:使用CHARRANGE时是以字节算,汉字算2个字节
    SendMessage txtTrace.hwnd, EM_EXGETSEL, 0, vSel
    vFind.chrg.cpMin = IIf(blnStart, 0, vSel.cpMax)
    vFind.chrg.cpMax = -1
    vFind.lpstrText = mstrFind
    k = SendMessage(txtTrace.hwnd, EM_FINDTEXT, FR_DOWN Or IIf(mblnMachCase, FR_MATCHCASE, 0), vFind)
    
    blnStart = False
    If k = -1 Then
        MsgBox "已找到文件尾部，下次查找将从头开始。", vbInformation, App.Title
        blnStart = True
    Else
        vSel.cpMin = k: vSel.cpMax = k + LenB(StrConv(mstrFind, vbFromUnicode))
        SendMessage txtTrace.hwnd, EM_EXSETSEL, 0, vSel
    End If
End Sub

Private Sub cmdFilter_Click()
    Dim blnShow As Boolean
    Dim i As Long, j As Long
    Dim lngRow As Long
    
    lngRow = -1
    
    With vsTrace
        For i = 0 To .Rows - 1 Step 5
            blnShow = True
            
            '仅全表扫描
            If chkFull.Value = 1 Then
                If .Cell(flexcpBackColor, i, 0) <> COLOR_FULL Then blnShow = False
            End If
            
            '命中率低于
            If IsNumeric(txtShoot.Text) And Val(txtShoot.Text) > 0 Then
                '为空的相当于没有命中率的概念,不显示
                If .TextMatrix(i + 4, COL_命中率) = "" Then
                    blnShow = False
                ElseIf Format(.TextMatrix(i + 4, COL_命中率), "0.0000") * 100 >= Val(txtShoot.Text) Then
                    blnShow = False
                End If
            End If
            
            For j = i To j + 4
                .RowHidden(j) = Not blnShow
            Next
            If blnShow And lngRow = -1 Then lngRow = i
        Next
        
        If lngRow <> -1 Then
            .Row = lngRow
            .ShowCell .Row, .Col
            .SetFocus
        Else
            txtSQL.Text = ""
            vsPlan.Rows = vsPlan.FixedRows
        End If
    End With
End Sub

Private Sub Form_Activate()
    RaiseEvent UpdateStatus(mstrFile & IIf(mstrSort <> "", "|排序:" & mstrSort, ""))
    Call Form_Resize '前一个窗口询问关闭后Resize有问题
End Sub

Private Sub Form_Deactivate()
    mfrmFind.Hide
End Sub

Private Sub Form_Load()
    Caption = gobjFile.GetFileName(mstrFile)
    

    '设置RTF不自动换行
    Call SendMessage(txtTrace.hwnd, EM_SETTARGETDEVICE, 0, 1)
    'txtTrace.RightMargin = 10000 '或这样也可以
    txtTrace.LoadFile mstrFile, rtfText

    
    mblnMultiRows = False
    Call FileToTable(True)
    
    Set mfrmFind = gfrmFind
    
    If gcnOracle <> "" Then
    
        gblnHasZltables = CheckTblExist("ZLTABLES")
    
        If mlngMinSize = 0 Then
            Call GetMidTabSize(mlngMinSize, mlngMaxSize)
        End If
    
        If mrsBigTbl Is Nothing Then
            Set mrsBigTbl = GetCheckObj(1, mlngMinSize, mlngMaxSize)
        End If
        
        If mrsBigIdx Is Nothing Then
            Set mrsBigIdx = GetCheckObj(2, mlngMinSize, mlngMaxSize)
        End If
        
        If mrsLowIdx Is Nothing Then
            Set mrsLowIdx = GetCheckObj(3, mlngMinSize, mlngMaxSize)
        End If
    End If
End Sub

Private Sub FileToTable(Optional ByVal blnInit As Boolean)
    Dim objText As TextStream
    Dim strLine As String, strTmp As String
    Dim strSql As String, lngFileRow As Long
    Dim strPlan As String, arrPlan As Variant
    Dim blnBegin As Boolean, lngCount As Long
    Dim intType As Integer, arrErr() As SQLError
    Dim i As Long, k As Long
    
    Screen.MousePointer = 11
    Me.Refresh
    
    On Error GoTo errH
    
    vsTrace.Rows = 0
    vsTrace.Redraw = flexRDNone
    Set objText = gobjFile.OpenTextFile(mstrFile, ForReading)
    
    '读取文件头部份
    Do While Not objText.AtEndOfStream
        strTmp = strLine
        strLine = objText.ReadLine
        
        'Trace文件分析时的排序方式
        If UCase(strLine) Like UCase("Sort options:") & "*" Then
            On Error Resume Next
            mstrSort = gcolSort("_" & UCase(Trim(Split(strLine, ":")(1))))
            If Err.Number <> 0 Then
                Err.Clear
                If UCase(Trim(Split(strLine, ":")(1))) = UCase("default") Then
                    mstrSort = "缺省"
                Else
                    mstrSort = Trim(Split(strLine, ":")(1))
                End If
            End If
            On Error GoTo errH
        End If
        
        '头部最后一行:*行
        If Replace(strLine, "*", "") = "" And strLine <> "" And UCase(Replace(strTmp, " ", "")) Like UCase("rows=*") Then
            Exit Do
        End If
    Loop
    If blnInit Then GoTo LineEnd
    
    '读取文件内容部份:各段以*行或空行结束(可能为多个空行,*行为大段结束,空行为小段结束)
    lngCount = 0 '总的SQL段个数
    intType = 0: blnBegin = False '每小段初始
    strSql = "": strPlan = "" '每大段初始
    Do While Not objText.AtEndOfStream
        strLine = objText.ReadLine
        
        If Replace(strLine, "*", "") = "" And strLine <> "" Then
            '本大段结束,重新初始化变量
            intType = 0: blnBegin = False
            strSql = "": strPlan = ""
            
        'chr(0)是解析大医二院(hp unix平台)的文件时发现的特殊字符
        ElseIf (strLine = "" Or strLine = Chr(0)) And blnBegin Then
            '段头行Data1存放执行计划内容
            If intType = 4 And strPlan <> "" Then
                With vsTrace
                    i = GetBaseRow(.Rows - 1)
                    .Cell(flexcpData, i, 1) = Mid(strPlan, 3)
                    
                    '全表扫描判断及颜色
                    arrPlan = Split(Mid(strPlan, 3), vbCrLf)
                    For k = 0 To UBound(arrPlan)
                        If InStr(arrPlan(k), "TABLE ACCESS FULL") > 0 _
                            And InStr(arrPlan(k), "TABLE ACCESS FULL DUAL") = 0 Or _
                            InStr(arrPlan(k), "INDEX FAST FULL SCAN") > 0 Or _
                            InStr(arrPlan(k), "INDEX FULL SCAN") > 0 Or _
                            InStr(arrPlan(k), "INDEX SKIP SCAN") > 0 _
                            Then Exit For
                    Next
                    If k <= UBound(arrPlan) Then
                        .Cell(flexcpBackColor, i, 0, .Rows - 1, .Cols - 1) = COLOR_FULL
                    End If
                End With
            End If
            
            If intType = 1 And strSql Like "SQL ID:*" And InStr(strSql, vbCrLf) = 0 Then
                '继续读取SQL文本
            Else
                '本小段结束,重新初始化变量
                intType = 0: blnBegin = False
            End If
        ElseIf strLine <> "" Then
            '到了汇总段,退出本循环,后面可以单独处理
            If UCase(strLine) = "OVERALL TOTALS FOR ALL NON-RECURSIVE STATEMENTS" Then Exit Do
            '到了结束段,退出本循环,后面可以单独处理
            If UCase(strLine) Like UCase("Trace file:*") Then Exit Do
            
            '错误语句段
            If UCase(strLine) = UCase("The following statements encountered a error during parse:") Then
                ReDim arrErr(0) '初始化数组
                Do While Not objText.AtEndOfStream
                    strLine = objText.ReadLine
                    
                    '不同错误SQL之间用---行间隔
                    If Replace(strLine, "-", "") = "" And strLine <> "" _
                        Or Replace(strLine, "*", "") = "" And strLine <> "" Then
                        arrErr(UBound(arrErr)).SQL = Trim(Split(Mid(strSql, 3), vbCrLf & "Error encountered:")(0))
                        arrErr(UBound(arrErr)).Err = Trim(Split(Mid(strSql, 3), vbCrLf & "Error encountered:")(1))
                        
                        If Replace(strLine, "-", "") = "" And strLine <> "" Then
                            strSql = ""
                            ReDim Preserve arrErr(UBound(arrErr) + 1)
                        End If
                    ElseIf Trim(strLine) <> "" Then
                        strSql = strSql & vbCrLf & strLine
                    End If
                    
                    '错误段结束,重新初始化变量
                    If Replace(strLine, "*", "") = "" And strLine <> "" Then
                        intType = 0: blnBegin = False
                        strSql = "": strPlan = "": GoTo LineNext
                    End If
                Loop
            End If
            
            '小段开始时,判断当前段的性质
            If Not blnBegin Then
                blnBegin = True '本段第一个非空行表示开始
                If strSql = "" Then
                    intType = 1 '大段开始strSQL才为空,并且开始就为SQL语句段
                ElseIf UCase(strLine) Like UCase("call*count*cpu*") Then
                    intType = 2 'Trace值表格段
                ElseIf UCase(strLine) Like UCase("Misses in library*") Then
                    intType = 3 'Trace说明段
                ElseIf UCase(strLine) Like UCase("Rows*Row Source Operation*") Then
                    intType = 4 '执行计划段
                ElseIf UCase(strLine) Like UCase("Elapsed times include*") Then
                    intType = 5 '等待时间段
                End If
            End If
            
            '根据各段的性质,作不同的处理
            If intType = 1 Then
                If strSql = "" Then lngFileRow = objText.Line - 1 - 1 '行号:1-n,读一行之后行指针后移,应-1,再-1是SQL前的空行
                strSql = IIf(strSql = "", "", strSql & vbCrLf) & strLine '段头行Data0存放SQL语句
            ElseIf intType = 2 Then
                '注：vsTrace.CellData存放的附加数据:
                '    段头行(RowData=SQLCount,Data0=SQL,Data1=Plan,Data2=Optimizer)
                '    每一行(Data of Cols-1=对应的源文件行数
                With vsTrace
                    If UCase(Left(strLine, 4)) = UCase("call") Then
                        lngCount = lngCount + 1
                        .AddItem Replace("(" & lngCount & ")|次数|CPU时间|总时间|物理读|一致读|当前读|记录数|命中率", "|", vbTab)
                        .RowData(.Rows - 1) = lngCount
                        .Cell(flexcpData, .Rows - 1, 0) = strSql
                        .CellBorderRange .Rows - 1, 0, .Rows - 1, .Cols - 1, &H808080, 0, 0, 0, 1, 0, 0
                        
                        '记录每大段对应的文件开始行号
                        .Cell(flexcpData, .Rows - 1, .Cols - 1) = lngFileRow
                    ElseIf UCase(Left(strLine, 5)) = UCase("Parse") Then
                        .AddItem Replace(ReplaceStr(strLine, " ", vbTab), "Parse", "解析")
                    ElseIf UCase(Left(strLine, 7)) = UCase("Execute") Then
                        .AddItem Replace(ReplaceStr(strLine, " ", vbTab), "Execute", "执行")
                    ElseIf UCase(Left(strLine, 5)) = UCase("Fetch") Then
                        .AddItem Replace(ReplaceStr(strLine, " ", vbTab), "Fetch", "提取")
                    ElseIf UCase(Left(strLine, 5)) = UCase("total") Then
                        .AddItem Replace(ReplaceStr(strLine, " ", vbTab), "total", "合计")
                        CalcAndShow命中率 .Rows - 1
                        .CellBorderRange .Rows - 1, 0, .Rows - 1, .Cols - 1, vbBlack, 0, 0, 0, 1, 0, 0
                    End If
                End With
            ElseIf intType = 3 Then
                If UCase(strLine) Like UCase("Optimizer goal:*") Then
                    i = GetBaseRow(vsTrace.Rows - 1)
                    vsTrace.Cell(flexcpData, i, 2) = Trim(Split(strLine, ":")(1)) '段头行Data2存放优化规则
                ElseIf UCase(strLine) Like UCase("Misses in library cache during parse:*") Then
                    With vsTrace
                        strTmp = Val(Split(strLine, ":")(1))
                        If strTmp <> "0" Then
                            .TextMatrix(.Rows - 4, COL_次数) = .TextMatrix(.Rows - 4, COL_次数) & ":" & strTmp
                        End If
                    End With
                ElseIf UCase(strLine) Like UCase("Misses in library cache during execute:*") Then
                    With vsTrace
                        strTmp = Val(Split(strLine, ":")(1))
                        If strTmp <> "0" Then
                            .TextMatrix(.Rows - 3, COL_次数) = .TextMatrix(.Rows - 3, COL_次数) & ":" & strTmp
                        End If
                    End With
                End If
            ElseIf intType = 4 Then
                If UCase(strLine) Like UCase("Rows*Row Source Operation*") Then
                    If mblnMultiRows = False Then
                        If strLine Like "Rows (1st) Rows (avg) Rows (max)*" Then mblnMultiRows = True
                    End If
                ElseIf strLine Like "-------*" Then
                Else
                    strPlan = strPlan & vbCrLf & strLine
                End If
            ElseIf intType = 5 Then
            End If
        End If
LineNext:
    Loop

LineEnd:
    '处理汇总段或结束段(strLine已读值)
    objText.Close
    vsTrace.AutoSize 0, vsTrace.Cols - 1
    If vsTrace.Rows > 0 Then vsTrace.Row = 0
    vsTrace.Redraw = flexRDDirect
    Screen.MousePointer = 0
    Exit Sub
errH:
    MsgBox Err.Number & ":" & vbCrLf & vbCrLf & Err.Description, vbCritical, App.Title
End Sub

Private Sub CalcAndShow命中率(ByVal lngRow As Long)
'参数：lngRow=合计行
'说明：命中率 = 1 - (物理读 / (逻辑读 = 一致读 + 当前读)),以Execute,Fetch行的数据合计为准
    Dim lng逻辑读 As Long, lng物理读 As Long
    Dim sng命中率 As Single
    
    With vsTrace
        lng物理读 = Val(.TextMatrix(lngRow - 1, COL_物理读)) + Val(.TextMatrix(lngRow - 2, COL_物理读))
        
        lng逻辑读 = Val(.TextMatrix(lngRow - 1, COL_一致读)) + Val(.TextMatrix(lngRow - 2, COL_一致读))
        lng逻辑读 = lng逻辑读 + Val(.TextMatrix(lngRow - 1, COL_当前读)) + Val(.TextMatrix(lngRow - 2, COL_当前读))
        
        If lng物理读 = 0 And lng逻辑读 = 0 Then
            '无此概念
            sng命中率 = -1
        ElseIf lng逻辑读 = 0 Then
            '命中率应为零
            sng命中率 = 0
        Else
            sng命中率 = 1 - lng物理读 / lng逻辑读
        End If
        
        If sng命中率 >= 0 Then
            .TextMatrix(lngRow, COL_命中率) = Format(sng命中率 * 100, "0.00") & "%"
        End If
    End With
End Sub

Private Function ReplaceStr(ByVal strText As String, ByVal strFrom As String, strTo As String) As String
    Do While InStr(strText, String(2, strFrom)) > 0
        strText = Replace(strText, String(2, strFrom), strFrom)
    Loop
    ReplaceStr = Replace(strText, strFrom, strTo)
End Function

Private Function GetBaseRow(ByVal lngRow As Long) As Long
    Dim i As Long
    
    GetBaseRow = -1
    
    With vsTrace
        If .RowData(lngRow) <> 0 Then
            GetBaseRow = lngRow: Exit Function
        Else
            For i = lngRow - 1 To .FixedRows Step -1
                If .RowData(i) > 0 Then
                    GetBaseRow = i: Exit Function
                End If
            Next
        End If
    End With
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 4 Then
        If MsgBox("确实要关闭当前窗口吗？", vbQuestion + vbYesNo + vbDefaultButton2, App.Title) = vbNo Then
            Cancel = 1: Exit Sub
        End If
    End If
End Sub

Private Sub Form_Resize()
    Dim sngH As Single, sngW As Single
    Dim lngFilter As Long
    
    If Me.WindowState = 1 Then Exit Sub
    
    On Error Resume Next
    
    If txtTrace.Visible Then
        Me.txtTrace.Left = 0
        Me.txtTrace.Top = 0
        Me.txtTrace.Width = Me.ScaleWidth
        Me.txtTrace.Height = Me.ScaleHeight
    Else
        sngH = txtSQL.Height / (txtSQL.Height + vsPlan.Height)
        sngW = vsTrace.Width / (vsTrace.Width + txtSQL.Width)
        
        fraFilter.Left = 0
        fraFilter.Top = 0
        fraFilter.Width = Me.ScaleWidth
        
        lngFilter = IIf(fraFilter.Visible, fraFilter.Height, 0)
        
        Line1.X1 = 0: Line1.X2 = fraFilter.Width
        Line1.Y1 = fraFilter.Height - 15: Line1.Y2 = Line1.Y1
        
        vsTrace.Left = 0
        vsTrace.Top = lngFilter
        vsTrace.Height = Me.ScaleHeight - lngFilter
        vsTrace.Width = (Me.ScaleWidth - fraLR.Width) * sngW
        
        fraLR.Top = lngFilter
        fraLR.Left = vsTrace.Left + vsTrace.Width
        fraLR.Height = Me.ScaleHeight - lngFilter
        
        txtSQL.Top = lngFilter
        txtSQL.Left = fraLR.Left + fraLR.Width
        txtSQL.Height = (Me.ScaleHeight - fraUD.Height - lngFilter) * sngH
        txtSQL.Width = Me.ScaleWidth - vsTrace.Width - fraLR.Width
        
        fraUD.Left = txtSQL.Left
        fraUD.Top = txtSQL.Top + txtSQL.Height
        fraUD.Width = txtSQL.Width
        
        vsPlan.Left = txtSQL.Left
        vsPlan.Top = fraUD.Top + fraUD.Height
        vsPlan.Width = txtSQL.Width
        vsPlan.Height = Me.ScaleHeight - txtSQL.Height - fraUD.Height - lngFilter
    End If
End Sub

Private Sub fraLR_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If vsTrace.Width + x < 2000 Or txtSQL.Width - x < 1000 Then Exit Sub
        
        fraLR.Left = fraLR.Left + x
        vsTrace.Width = vsTrace.Width + x
        
        txtSQL.Left = txtSQL.Left + x
        txtSQL.Width = txtSQL.Width - x
        
        fraUD.Left = fraUD.Left + x
        fraUD.Width = fraUD.Width - x
        
        vsPlan.Left = vsPlan.Left + x
        vsPlan.Width = vsPlan.Width - x
    End If
End Sub

Private Sub fraUD_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If txtSQL.Height + y < 1000 Or vsPlan.Height - y < 1000 Then Exit Sub
        
        fraUD.Top = fraUD.Top + y
        txtSQL.Height = txtSQL.Height + y
        vsPlan.Top = vsPlan.Top + y
        vsPlan.Height = vsPlan.Height - y
    End If
End Sub

Private Sub mfrmFind_Find(ByVal Text As String, ByVal MatchCase As Boolean)
    If Not frmMain.ActiveForm Is Me Then Exit Sub
    
    mstrFind = Text
    mblnMachCase = MatchCase
    Call SearchText
End Sub

Private Sub txtShoot_GotFocus()
    txtShoot.SelStart = 0: txtShoot.SelLength = Len(txtShoot.Text)
End Sub

Private Sub txtShoot_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtSQL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        txtSQL.SelStart = 0: txtSQL.SelLength = Len(txtSQL.Text)
    End If
End Sub

Private Sub txtTrace_SelChange()
    Dim lngLine As Long
    
    '用GetLineFromChar很慢
    lngLine = SendMessage(txtTrace.hwnd, EM_LINEINDEX, -1, 0)
    lngLine = SendMessage(txtTrace.hwnd, EM_LINEFROMCHAR, lngLine, 0) + 1
    RaiseEvent UpdateStatus(mstrFile & IIf(mstrSort <> "", "|排序:" & mstrSort, "|") & "|行号:" & lngLine)
End Sub

Private Sub vsTrace_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strPlan As String, arrPlan As Variant
    Dim lngRow As Long, strRow As String
    Dim lngRowAvg As Long, lngRowMax As Long
    Dim strOpti As String, i As Long, lngTop As Long
    
    
    If NewRow <> OldRow And NewRow <> -1 Then
        With vsTrace
            
            If .Rows = 0 Then Exit Sub
            
            i = GetBaseRow(NewRow)
            If txtSQL.Text <> .Cell(flexcpData, i, 0) Then
                txtSQL.Text = .Cell(flexcpData, i, 0)
            End If
            strPlan = .Cell(flexcpData, i, 1)
            strOpti = .Cell(flexcpData, i, 2)
            
            RaiseEvent UpdateStatus(mstrFile & IIf(mstrSort <> "", "|排序:" & mstrSort, "|") & "|行号:" & .Cell(flexcpData, i, .Cols - 1))
        End With
        
        With vsPlan
            .Redraw = flexRDNone
            If mblnMultiRows And .Cols = 2 Then
                '11G解析出的执行计划包含这几列：Rows (1st) Rows (avg) Rows (max)  Row Source Operation
                
                .Cols = 4
                .TextMatrix(0, 0) = "Rows (1st)"
                .TextMatrix(0, 1) = "Rows (avg)"
                .ColAlignment(1) = flexAlignRightCenter
                .TextMatrix(0, 2) = "Rows (max)"
                .TextMatrix(0, 3) = "Row Source Operation"
                .OutlineCol = 3
            ElseIf mblnMultiRows = False And .Cols = 4 Then
                .Cols = 2
                .TextMatrix(0, 0) = "行数"
                .TextMatrix(0, 1) = "操作" & IIf(strOpti <> "", "(优化 = " & strOpti & ")", "")
                .ColAlignment(1) = flexAlignLeftCenter
                .OutlineCol = 1
            End If
            
            .Rows = .FixedRows
            .FixedAlignment(1) = flexAlignLeftCenter
            
            If strPlan <> "" Then
                arrPlan = Split(strPlan, vbCrLf)
                For i = 0 To UBound(arrPlan)
                    If mblnMultiRows Then
                        lngRow = Val(Mid(arrPlan(i), 1, 10))
                        lngRowAvg = Val(Mid(arrPlan(i), 12, 10))
                        lngRowMax = Val(Mid(arrPlan(i), 22, 10))
                        strRow = Mid(arrPlan(i), 35)
                        .AddItem lngRow & vbTab & lngRowAvg & vbTab & lngRowMax & vbTab & Trim(Split(strRow, "(object id")(0))
                    Else
                        lngRow = Val(Trim(arrPlan(i)))
                        strRow = Mid(Trim(arrPlan(i)), InStr(Trim(arrPlan(i)), " ") + 2)
                        .AddItem lngRow & vbTab & Trim(Split(strRow, "(object id")(0))
                    End If
                    
                    .RowOutlineLevel(.Rows - 1) = Len(strRow) - Len(LTrim(strRow))
                    .IsSubtotal(.Rows - 1) = True
                    
                    If InStr(arrPlan(i), "TABLE ACCESS FULL") > 0 _
                        And InStr(arrPlan(i), "TABLE ACCESS FULL DUAL") = 0 Or _
                            InStr(arrPlan(i), "INDEX FAST FULL SCAN") > 0 Or _
                            InStr(arrPlan(i), "INDEX FULL SCAN") > 0 Or _
                            InStr(arrPlan(i), "INDEX SKIP SCAN") > 0 _
                        Then
                        .Cell(flexcpBackColor, .Rows - 1, 0, .Rows - 1, .Cols - 1) = COLOR_FULL
                        If lngTop = 0 Then lngTop = .Rows - 1
                    End If
                Next
                .Row = .FixedRows
            End If
            .CellBorderRange 0, 0, .Rows - 1, 0, &H808080, 0, 0, 1, 0, 0, 0
            .CellBorderRange .FixedRows - 1, 0, .FixedRows - 1, .Cols - 1, &H808080, 0, 0, 0, 1, 1, 0
            .CellBorderRange .Rows - 1, 0, .Rows - 1, .Cols - 1, &H808080, 0, 0, 0, 1, 1, 0
            .AutoSize 0, .Cols - 1
            .Redraw = flexRDDirect
            
            .TopRow = lngTop
        End With
        
        Call CheckSqlPlan(vsPlan, 1, 1, mrsBigTbl, mrsBigIdx, mrsLowIdx)
    End If
End Sub
