VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPartogramInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "产程信息录入"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6360
   Icon            =   "frmPartogramInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame fraLine 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   3615
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "帮助(&H)"
      Height          =   350
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancle 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   4920
      TabIndex        =   11
      Top             =   3840
      Width           =   1100
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   3720
      TabIndex        =   10
      Top             =   3840
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid VsfData 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6225
      _cx             =   10980
      _cy             =   6376
      Appearance      =   0
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
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   4
      Cols            =   4
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   255
      RowHeightMax    =   5000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPartogramInfo.frx":000C
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
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   -1  'True
      DataMember      =   ""
      ComboSearch     =   3
      AutoSizeMouse   =   0   'False
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
      Begin MSComCtl2.MonthView mthView 
         Height          =   2220
         Left            =   720
         TabIndex        =   5
         Top             =   2160
         Visible         =   0   'False
         Width           =   2805
         _ExtentX        =   4948
         _ExtentY        =   3916
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   181272577
         CurrentDate     =   40899
      End
      Begin VB.PictureBox picInput 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         ScaleHeight     =   225
         ScaleWidth      =   1425
         TabIndex        =   1
         Top             =   1800
         Visible         =   0   'False
         Width           =   1455
         Begin MSComCtl2.UpDown UD 
            Height          =   300
            Left            =   241
            TabIndex        =   3
            Top             =   0
            Visible         =   0   'False
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            BuddyControl    =   "txtInput"
            BuddyDispid     =   196618
            OrigLeft        =   120
            OrigRight       =   375
            OrigBottom      =   255
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.CommandButton cmdDown 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   480
            Picture         =   "frmPartogramInfo.frx":006E
            Style           =   1  'Graphical
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   0
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.TextBox txtInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   0
            Width           =   240
         End
      End
      Begin VB.ListBox lstSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   2550
         Index           =   0
         ItemData        =   "frmPartogramInfo.frx":03B0
         Left            =   3600
         List            =   "frmPartogramInfo.frx":03C6
         TabIndex        =   6
         Top             =   720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ListBox lstSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   2550
         Index           =   1
         ItemData        =   "frmPartogramInfo.frx":03FE
         Left            =   4530
         List            =   "frmPartogramInfo.frx":0414
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   1485
      End
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPartogramInfo.frx":044C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "提示:"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   1680
      TabIndex        =   8
      Top             =   3840
      Width           =   450
   End
End
Attribute VB_Name = "frmPartogramInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum ColEnum
    COL_NULL
    COL_Group
    COL_Name
    COL_value
End Enum

Const mlngColEditor As Long = 3
Private mlngFileID As Long  '文件ID
Private mlngFileIndex As Long '文件份数
Private mlngFileFormatID As Long '文件格式ID
Private mrsPartogram As New ADODB.Recordset '可添加的产程要素信息
Private mblnOK As Boolean
Private mblnInit As Boolean
Private mintFace As Integer '要素表示 0-文本;1-上下;2-下拉;3-复选;4-单选
Private mintType As String   '要素类型 0-数值;1-文本;2-日期;3-逻辑
Private mblnShow As Boolean
Private mblnBlowup As Boolean
Private mblnStart As Boolean
Private mblnChange As Boolean


Public Function ShowMe(ByVal frmParent As Form, ByVal lngFileID As Long, ByVal lngFileIndex As Long, ByVal lngFileFormatID As Long, ByVal rsPartogram As ADODB.Recordset, _
    Optional ByVal bytSize As Byte = 0) As Boolean
 '------------------------------------------------------
 '功能：完成产程数据录入
 '参数：frmParent :调用窗体对象
 '      lngFileID :文件ID
 '      lngFileIndex :第几个婴儿
 '      lngFileFormatID :文件格式ID
 '      rsPartogram ：可添加的表上、表下标签内容（中文名,替换域,类型,长度,小数,单位,表示法,数值域,必填）
 '------------------------------------------------------
    mblnOK = False
    mlngFileID = lngFileID
    mlngFileIndex = lngFileIndex
    mlngFileFormatID = lngFileFormatID
    Set mrsPartogram = rsPartogram
    mblnStart = True
    mblnBlowup = (bytSize = 1) '(zlDatabase.GetPara("护理文件显示模式", glngSys, 1255, 0) = 1)
    Me.FontSize = IIf(bytSize = 0, 9, IIf(bytSize = 1, 12, bytSize))
    
    If Not zlRefresh Then Exit Function

    Me.Show vbModal, frmParent

    ShowMe = mblnOK
End Function

Private Function zlRefresh() As Boolean
'---------------------------------------
'刷新数据信息
'---------------------------------------
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand

    mblnInit = False
    mblnShow = False
    mblnChange = False

    Call InitCons
    '提取表上表下内容
    gstrSQL = _
        " Select '' 空, '表上标签' 分组名, d.要素名称, a.内容" & vbNewLine & _
        " From 病历文件结构 D, 病历文件结构 P, 产程要素内容 A" & vbNewLine & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表上标签' And d.要素名称 = a.名称(+) And a.文件id(+) = [2] And" & vbNewLine & _
        "      a.婴儿(+) = [3]" & vbNewLine & _
        " Union All " & vbNewLine & _
        " Select '' 空,'表下标签' 分组名, d.要素名称, a.内容" & vbNewLine & _
        " From 病历文件结构 D, 病历文件结构 P, 产程要素内容 A" & vbNewLine & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表下标签' And d.要素名称 = a.名称(+) And a.文件id(+) = [2] And" & vbNewLine & _
        "      a.婴儿(+) = [3]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取要素信息", mlngFileFormatID, mlngFileID, mlngFileIndex)
    
    Call InitTabFormat(rsTemp)
    mblnInit = True
    Call VsfData_AfterRowColChange(0, COL_Name, VsfData.FixedRows, COL_value)
    zlRefresh = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub InitCons()
    '隐藏输入控件
    mintType = -1
    mintFace = -1
    picInput.Visible = False
    lstSelect(0).Visible = False
    lstSelect(1).Visible = False
    mthView.Visible = False
    UD.Visible = False
    cmdDown.Visible = False
End Sub

Private Sub InitTabFormat(ByVal rsTemp As ADODB.Recordset)
    Dim i As Integer, j As Integer
    With VsfData
        .Cols = 4
        .Rows = 2
        .FixedCols = 3
        .FixedRows = 1
        .MergeCells = flexMergeFixedOnly
        .MergeCol(COL_Group) = True
        .TextMatrix(0, COL_NULL) = ""
        .ColWidth(COL_NULL) = 255
        .TextMatrix(0, COL_Group) = "分组名"
        .ColWidth(COL_Group) = 1000
        .TextMatrix(0, COL_Name) = "要素名称"
        .ColWidth(COL_Name) = 1500
        .TextMatrix(0, COL_value) = "要素内容"
        .ColWidth(COL_value) = 3400
        .RowHeightMin = 300
        .FontSize = 9 + 9 * IIf(mblnBlowup = True, 1, 0) / 3
        '完成数据绑定
        If rsTemp.RecordCount = 0 Then
            .Rows = 2
        Else
            rsTemp.MoveFirst
            mrsPartogram.Filter = 0
            j = .FixedRows
            For i = 1 To rsTemp.RecordCount
                mrsPartogram.Filter = "中文名='" & NVL(rsTemp!要素名称) & "' And 替换域<>1"
                If mrsPartogram.RecordCount > 0 Then
                    If .Rows <= j Then .Rows = .Rows + 1
                    .TextMatrix(j, COL_NULL) = NVL(rsTemp!空)
                    .TextMatrix(j, COL_Group) = NVL(rsTemp!分组名)
                    .TextMatrix(j, COL_Name) = Replace(NVL(rsTemp!要素名称), ";", "")
                    .TextMatrix(j, COL_value) = Replace(NVL(rsTemp!内容), "[ZLSOFTLPF]", "")
                    .MergeRow(j) = True
                    j = j + 1
                End If
                If i < rsTemp.RecordCount Then rsTemp.MoveNext
            Next i
        End If

        .COL = COL_value: .ROW = 1
        .Cell(flexcpAlignment, 0, 0, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .FocusRect = flexFocusSolid

        .AutoResize = True
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 0, .Cols - 1
        .AutoResize = False
    End With
End Sub

Private Sub AdjustRowFlag(ByRef objVsf As Object, ByVal intRow As Integer)
    '-----------------------------------------------------------------------------------------
    '功能:
    '参数:
    '-----------------------------------------------------------------------------------------
    If objVsf.FixedCols = 0 Then Exit Sub

    If Not (objVsf.Cell(flexcpPicture, intRow, COL_NULL) Is Nothing) Then Exit Sub
    Set objVsf.Cell(flexcpPicture, 1, COL_NULL, objVsf.Rows - 1, COL_NULL) = Nothing
    Set objVsf.Cell(flexcpPicture, intRow, COL_NULL) = ils16.ListImages(1).Picture
End Sub


Private Sub cmdCancle_Click()
    mblnChange = False
    mblnOK = False
    Unload Me
End Sub

Private Sub cmdDown_Click()
    Dim arrData
    Dim i As Integer, j As Integer
    Dim strName As String, strBound As String
    Dim intIndex As Integer
    Dim CellRect As RECT
    Dim strValue As String

    If mblnShow = False Or mintFace = -1 Then Exit Sub

        CellRect.Left = picInput.Left
        CellRect.Top = picInput.Top + picInput.Height
        CellRect.Bottom = VsfData.CellHeight
        CellRect.Right = VsfData.CellWidth
    strValue = Trim(txtInput.Text)
    If mintType = 2 Then '日期下拉
        With mthView
            If IsDate(strValue) Then
                .Value = Format(strValue, "YYYY-MM-DD")
            Else
                .Value = Format(zlDatabase.Currentdate, "YYYY-MM-DD")
            End If
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Font.Name = VsfData.FontName
            .Font.Size = VsfData.FontSize
            If .Height + .Top > VsfData.Height Then
                .Top = VsfData.Height - .Height
            End If
            If .Width < CellRect.Right Then
                .Left = CellRect.Right + CellRect.Left - .Width
            End If
            .Visible = True
            .ZOrder 0
            .SetFocus
        End With
    Else '文本下拉
        strName = VsfData.TextMatrix(VsfData.ROW, COL_Name)
        mrsPartogram.Filter = "中文名='" & strName & "'"
        strBound = NVL(mrsPartogram!数值域)
        If Left(strBound, 1) = ";" Then strBound = Mid(strBound, 2)
        If strBound <> "" Then strBound = ";" & strBound
        intIndex = 0
        lstSelect(intIndex).Clear
        arrData = Split(strBound, ";")
        j = UBound(arrData)
        lstSelect(intIndex).AddItem 0 & "-"
        lstSelect(intIndex).ListIndex = 0
        For i = 1 To j
            If arrData(i) <> "" Then
                If Mid(arrData(i), 1, 1) = "√" Then
                    lstSelect(intIndex).AddItem i & "-" & Mid(arrData(i), 2)
                    lstSelect(intIndex).ListIndex = i
                Else
                    lstSelect(intIndex).AddItem i & "-" & arrData(i)
                End If
            End If
        Next
        
        If strValue <> "" Then
            strValue = Replace(strValue, vbCrLf, "")
            For i = 0 To j
                If InStr(1, "," & strValue & ",", "," & Mid(lstSelect(intIndex).List(i), InStr(1, lstSelect(intIndex).List(i), "-") + 1) & ",") <> 0 Then
                    lstSelect(intIndex).Selected(i) = True
                End If
            Next
        End If
        '显示
        With lstSelect(intIndex)
            .Left = CellRect.Left
            .Top = CellRect.Top
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Height = .ListCount * 300
            If .Height < CellRect.Bottom Then .Height = CellRect.Bottom
            .Width = LenB(StrConv(.List(.ListCount \ 2), vbFromUnicode)) * 120 + 500    '以中间项的长度为依据
            If .Width < CellRect.Right Then .Width = CellRect.Right
            If .Height + .Top > VsfData.Height Then
                .Top = VsfData.Height - .Height
            End If
            .Visible = True
            .ZOrder 0
            .SetFocus
        End With
    End If
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOK_Click()
    Dim strPara As String, strSQL As String
    Dim strName As String, strValue As String
    Dim intRow As Integer
    If mblnChange = True Then
        For intRow = VsfData.FixedRows To VsfData.Rows - 1
            If InStr(1, strName, "[ZLSOFTLPF]" & VsfData.TextMatrix(intRow, COL_Name) & "[ZLSOFTLPF]") = 0 And Trim(VsfData.TextMatrix(intRow, COL_value)) <> "" Then
                strName = strName & "[ZLSOFTLPF]" & VsfData.TextMatrix(intRow, COL_Name) & "[ZLSOFTLPF]"
                strPara = strPara & "[ZLSOFTLPF]" & VsfData.TextMatrix(intRow, COL_Name) & ";" & VsfData.TextMatrix(intRow, COL_value)
            End If
        Next intRow
        If Left(strPara, 11) = "[ZLSOFTLPF]" Then strPara = Mid(strPara, 12)
        '保存数据信息
        strSQL = "ZL_产程要素内容_UPDATE("
        '文件ID_IN IN 产程要素内容.文件ID%TYPE,
        strSQL = strSQL & mlngFileID & ","
        '婴儿_IN IN  产程要素内容.婴儿 %TYPE,
        strSQL = strSQL & mlngFileIndex & ",'"
        'strPara IN Varchar2 --参数格式为：要素名称;要素内容|要素名称;要素内容
        strSQL = strSQL & strPara & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, "ZL_产程要素内容_UPDATE")
        mblnChange = False
        mblnOK = True
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    If Not mblnStart Then Exit Sub
    Me.Width = Me.Width + Me.Width * IIf(mblnBlowup = True, 1, 0) / 3
    Me.Height = Me.Height + Me.Height * IIf(mblnBlowup = True, 1, 0) / 3
    Me.FontSize = Me.FontSize + Me.FontSize * IIf(mblnBlowup = True, 1, 0) / 3
    lblInfo.FontSize = Me.FontSize
    mblnStart = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    '以下字符做为数据分隔符或更新记录集的分隔符，因此不允许录入
    If KeyAscii = 39 Or KeyAscii = 13 Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyEscape And mblnShow Then
        mblnShow = False
        Call InitCons
    End If
End Sub

Private Sub Form_Resize()

    With cmdHelp
        .Left = 120
        .Top = Me.ScaleHeight - .Height - 120
    End With

    With cmdCancle
        .Top = cmdHelp.Top
        .Left = Me.ScaleWidth - .Width - 120
    End With

    With cmdOK
        .Top = cmdCancle.Top
        .Left = cmdCancle.Left - .Width - 120
    End With

    With fraLine
        .Left = 60
        .Top = cmdOK.Top - 60
        .Width = Me.ScaleWidth - 120
    End With

    With lblInfo
        .Left = 120
        .Top = fraLine.Top - TextHeight("刘") - 60
    End With

    With VsfData
        .Left = 0
        .Top = 0
        .Width = Me.ScaleWidth
        .Height = lblInfo.Top - 60
    End With
End Sub

Private Sub lstSelect_DblClick(Index As Integer)
    Call lstSelect_KeyDown(Index, vbKeyReturn, 0)
End Sub

Private Sub lstSelect_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim strText As String
    Dim intIndex As Integer
    If KeyCode = vbKeyReturn Then
        If mintFace = 2 Then '文本下拉
            If InStr(1, lstSelect(Index).Text, "-") <> 0 Then
                strText = Split(lstSelect(Index).Text, "-")(1)
            Else
                strText = ""
            End If
            txtInput.Text = strText
            lstSelect(Index).Visible = False
            If picInput.Visible = True Then picInput.SetFocus
        Else
            Call MoveNextCell
        End If
    ElseIf (KeyCode = vbKeyDown Or KeyCode = vbKeyUp) And Shift = vbShiftMask Then
        Call MoveNextCell(KeyCode, 0)
    End If
    
End Sub

Private Sub mthView_DateDblClick(ByVal DateDblClicked As Date)
    txtInput.Text = Format(DateDblClicked, "YYYY-MM-DD")
    mthView.Visible = False
    If picInput.Visible = True Then picInput.SetFocus
End Sub

Private Sub mthView_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If IsDate(mthView.Value) Then Call mthView_DateDblClick(CDate(mthView.Value))
    End If
End Sub

Private Sub picInput_GotFocus()
    If txtInput.Visible = True Then
        txtInput.SetFocus
    End If
End Sub

Private Sub picInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift <> vbShiftMask Then
        If KeyCode = vbKeyReturn Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Then
            Call MoveNextCell(KeyCode, Shift)
        End If
    Else
        If KeyCode = vbKeyDown And mintFace = 2 Then
            Call cmdDown_Click
        End If
    End If
End Sub

Private Sub txtInput_GotFocus()
    txtInput.SelStart = 0
    txtInput.SelLength = Len(txtInput.Text)
    If lstSelect(0).Visible = True And ((mintFace = 2 And mintType = 1) Or mintFace = 4) Then
        lstSelect(0).SetFocus
    ElseIf mthView.Visible = True And mintFace = 2 And mintType = 2 Then
        mthView.SetFocus
    ElseIf lstSelect(1).Visible = True And mintFace = 3 Then
        lstSelect(1).SetFocus
    End If
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    Call picInput_KeyDown(KeyCode, Shift)
End Sub

Private Sub VsfData_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strName As String, strBound As String, strInfo As String
    
    If mblnInit = False Then Exit Sub
    If OldRow = NewRow Then Exit Sub
    Call AdjustRowFlag(VsfData, NewRow)
    strName = VsfData.TextMatrix(NewRow, COL_Name)
    If mrsPartogram Is Nothing Then Exit Sub
    
    '显示产程项目值域信息
    mrsPartogram.Filter = 0
    mrsPartogram.Filter = "中文名='" & strName & "'"
    If mrsPartogram.RecordCount > 0 Then
        strBound = NVL(mrsPartogram!数值域, "")
        If Left(strBound, 1) = ";" Then strBound = Mid(strBound, 2)
        If strBound <> "" Then
            If Val(NVL(mrsPartogram!类型, 0)) = 0 Then
                strInfo = "数值域:" & Split(strBound, ";")(0) & "～" & Split(strBound, ";")(1)
            Else
                strInfo = "数值域;" & strBound
            End If
        End If
        If Val(NVL(mrsPartogram!表示法, 0)) = 2 Then
            strInfo = strInfo & IIf(strInfo = "", "", Space(2)) & "[按SHIFT+↓弹出下拉框]"
        End If
    End If
    
    lblInfo.Caption = "提示：" & strInfo
    lblInfo.Tag = lblInfo.Caption
End Sub

Private Sub VsfData_DblClick()
    Call vsfdata_KeyDown(Asc("A"), 0)
End Sub

Private Sub VsfData_EnterCell()

    '隐藏以显示的控件
    Select Case mintFace
        Case 0, 1, 2
            picInput.Visible = False
            If mintFace = 2 Then
                lstSelect(0).Visible = False
            End If
        Case 3
            lstSelect(1).Visible = False
        Case 4
            lstSelect(0).Visible = False
    End Select
    mthView.Visible = False
    UD.Visible = False
    cmdDown.Visible = False

    mintType = -1: mintFace = -1
    If mblnShow = False Or VsfData.COL <> COL_value Then Exit Sub

    Call ShowInput

    '获取焦点
    Select Case mintFace
        Case 0, 1, 2
            picInput.SetFocus
        Case 3
            lstSelect(1).SetFocus
        Case 4
            lstSelect(0).SetFocus
    End Select
End Sub

Private Sub vsfdata_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngStart As Long
    Dim intLevel As Integer
    Dim strField As String, strKey As String, strValue As String
    On Error GoTo errHand

    If KeyCode = vbKeyReturn Then
        Call MoveNextCell
    Else
        If Not (KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or Shift <> 0) Then
            mblnShow = True
            Call VsfData_EnterCell
        End If
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function ShowInput(Optional ByVal intRow As Integer = -1) As Boolean
'显示相应的 编辑控件
    Dim CellRect As RECT
    Dim arrData
    Dim intIndex As Integer, i As Integer, j As Integer, k As Integer
    Dim strItemName As String, strBound As String, strLen As String, strValue As String

    If intRow = -1 Then intRow = VsfData.ROW
    CellRect.Left = VsfData.CellLeft + VsfData.Left
    CellRect.Top = VsfData.CellTop + VsfData.Top
    CellRect.Bottom = VsfData.CellHeight + 0
    CellRect.Right = VsfData.CellWidth + 0

    mintType = -1
    mintFace = -1
    strItemName = VsfData.TextMatrix(intRow, COL_Name)
    strValue = VsfData.TextMatrix(intRow, COL_value)
    '中文名,替换域,类型,长度,小数,单位,表示法,数值域,必填
    mrsPartogram.Filter = 0
    mrsPartogram.Filter = "中文名='" & strItemName & "'"
    '确定项目类型
    mintFace = Val(NVL(mrsPartogram!表示法, 0))
    mintType = Val(NVL(mrsPartogram!类型, 0))
    strBound = NVL(mrsPartogram!数值域)
    If Left(strBound, 1) = ";" Then strBound = Mid(strBound, 2)
    strLen = Val(NVL(mrsPartogram!长度)) & ";" & Val(NVL(mrsPartogram!小数))
    '类型为逻辑处理为 文件下拉
    If mintType = 3 Then mintFace = 2: mintType = 1
    Select Case mintFace
    Case 0, 1, 2
        With picInput
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Width = CellRect.Right
            .Height = CellRect.Bottom
            If .Height + .Top > VsfData.Height Then
                .Top = VsfData.Height - .Height
            End If
            .Visible = True
        End With
        '文本或数字项目
        txtInput.Visible = True
        If Val(strLen) <> 0 Then
            txtInput.MaxLength = Val(Split(strLen, ";")(0)) + IIf(Val(Split(strLen, ";")(1)) = 0, 0, 1) '小数位数要加上小数点
        Else
            txtInput.MaxLength = 0
        End If

        With txtInput
            .Top = 0
            .Text = strValue
            .Width = CellRect.Right
            .Height = CellRect.Bottom
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Tag = .Text
            If mintFace = 1 Then
                arrData = Split(strBound, ";")
                UD.Top = 0
                .Width = .Width - UD.Width
                UD.Left = .Width
                UD.Height = .Height
                UD.Min = 0: UD.Max = 10
                UD.Increment = 1
                If UBound(arrData) > 0 Then
                    UD.Min = Val(arrData(0))
                    UD.Max = Val(arrData(1))
                End If
                UD.Visible = True
            ElseIf mintFace = 2 Then
                cmdDown.Top = 0
                .Width = .Width - cmdDown.Width
                cmdDown.Left = .Width
                cmdDown.Height = .Height
                cmdDown.Visible = True
            End If
            .Width = .Width - (180 + IIf(mblnBlowup, 180 * 1 / 3, 0)) / 2 '宋体9号时减去90,字体越大扣除的边距越小,以保证文本框分行与实际一致
        End With
    Case 3, 4
        intIndex = IIf(mintFace = 3, 1, 0)
        '加载数据
        lstSelect(intIndex).Clear
        If Left(strBound, 1) = ";" Then strBound = Mid(strBound, 2)
        If strBound <> "" Then strBound = ";" & strBound
        k = 1
        If intIndex = 0 Then
            lstSelect(intIndex).AddItem 0 & "-"
            lstSelect(intIndex).ListIndex = 0
        End If
        arrData = Split(strBound, ";")
        j = UBound(arrData)
        For i = k To j
            If arrData(i) <> "" Then
                If Mid(arrData(i), 1, 1) = "√" Then
                    lstSelect(intIndex).AddItem lstSelect(intIndex).NewIndex + 1 & "-" & Mid(arrData(i), 2)
                    lstSelect(intIndex).ListIndex = lstSelect(intIndex).NewIndex
                Else
                    lstSelect(intIndex).AddItem lstSelect(intIndex).NewIndex + 1 & "-" & arrData(i)
                End If
            End If
        Next
        '多选且已录入数据的情况下
        If strValue <> "" Then
            strValue = Replace(strValue, vbCrLf, "")
            For i = 0 To j
                If InStr(1, "," & strValue & ",", "," & Mid(lstSelect(intIndex).List(i), InStr(1, lstSelect(intIndex).List(i), "-") + 1) & ",") <> 0 Then
                    lstSelect(intIndex).Selected(i) = True
                End If
            Next
        End If
        '显示
        With lstSelect(intIndex)
            .Left = CellRect.Left
            .Top = CellRect.Top
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Height = .ListCount * 300
            If .Height < CellRect.Bottom Then .Height = CellRect.Bottom
            .Width = LenB(StrConv(.List(.ListCount \ 2), vbFromUnicode)) * 120 + 500    '以中间项的长度为依据
            If .Width < CellRect.Right Then .Width = CellRect.Right
            If .Height + .Top > VsfData.Height Then
                .Top = VsfData.Height - .Height
            End If
            .Visible = True
            .Tag = strValue
        End With
    End Select

    ShowInput = True
End Function

Private Sub MoveNextCell(Optional KeyCode As Integer = vbKeyReturn, Optional Shift As Integer = 0)
'进行数据校验和单元格移动
    Dim intRow As Integer
    Dim strRetrun As String, strErrMsg As String
    Dim blnShow As Boolean
    
    If mintFace >= 0 And Shift = vbShiftMask And (KeyCode = vbKeyUp Or KeyCode = vbKeyDown) Then Exit Sub
    If mintFace >= 0 And KeyCode = vbKeyReturn Then
        '完成数据校验和保存
        If Not CheckInput(strRetrun, strErrMsg) Then
            lblInfo.Caption = "提示：" & strErrMsg
            Exit Sub
        Else
            lblInfo.Caption = lblInfo.Tag
        End If
        '完成赋值工作
        VsfData.TextMatrix(VsfData.ROW, COL_value) = Replace(strRetrun, "[ZLSOFTLPF]", "")
    End If
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyDown Then
toMoveNextCol:
        If VsfData.COL < mlngColEditor Then
            VsfData.COL = VsfData.COL + 1
            If VsfData.ColWidth(VsfData.COL) = 0 Or VsfData.ColHidden(VsfData.COL) Then GoTo toMoveNextCol
        Else
toMoveNextRow:
            '跳到下一行
            intRow = 1
            If VsfData.ROW + intRow < VsfData.Rows Then
                VsfData.ROW = VsfData.ROW + intRow
            Else
                blnShow = mblnShow
                mblnShow = False
                Call VsfData_EnterCell
                mblnShow = blnShow
            End If
            If VsfData.RowHidden(VsfData.ROW) Then GoTo toMoveNextRow
            VsfData.COL = COL_value
        End If
    Else
toMovePrevCol:
        If VsfData.COL > mlngColEditor Then
            VsfData.COL = VsfData.COL - 1
            If VsfData.ColWidth(VsfData.COL) = 0 Or VsfData.ColHidden(VsfData.COL) Then GoTo toMovePrevCol
        Else
toMovePrevRow:
'            '跳到上一行
            intRow = 1
            If VsfData.ROW > VsfData.FixedRows Then
                VsfData.ROW = VsfData.ROW - intRow
            End If
            If VsfData.RowHidden(VsfData.ROW) Then GoTo toMovePrevRow
            VsfData.COL = COL_value
        End If
    End If
    If VsfData.ColIsVisible(VsfData.COL) = False Then
        VsfData.LeftCol = VsfData.COL
    End If
    If VsfData.RowIsVisible(VsfData.ROW) = False Then
        VsfData.TopRow = VsfData.ROW
    End If
End Sub

Private Function CheckInput(strReturn As String, strInfo As String) As Boolean
    Dim i As Integer, j As Integer
    Dim strText As String, strOldText As String
    Dim intIndex As Integer
    Dim arrDate
    '检查录入数据的合法性(中文也认为是一个字符,考虑到体温项目等存在不升\外出等信息)
    '返回的数据,如果一列绑定多个项目,以单引号做为分隔符

    'mintType:0=文本框录入;1=单选;2=多选;3=选择;4-血压或一列绑定了两个项目,其格式类似血压的输入项目;5=一列绑定了两个项目且均是选择项目;
    '6=一列绑定N个项目,手工录入
    Select Case mintFace
    Case 0, 1, 2
        strText = txtInput.Text
        strOldText = txtInput.Tag
    Case 3, 4   '免检
        intIndex = IIf(mintFace = 3, 2, 1)
        If mintFace = 4 Then
            If InStr(1, lstSelect(intIndex - 1).Text, "-") <> 0 Then
                strText = Split(lstSelect(intIndex - 1).Text, "-")(1)
            Else
                strText = ""
            End If
        Else
            j = lstSelect(intIndex - 1).ListCount
            For i = 1 To j
                If lstSelect(intIndex - 1).Selected(i - 1) Then
                    strText = strText & "," & Split(lstSelect(intIndex - 1).List(i - 1), "-")(1)
                End If
            Next
            If strText <> "" Then strText = Mid(strText, 2)
        End If
        strOldText = lstSelect(intIndex - 1).Tag
    End Select

    If mintType = 0 Or (mintType = 1 And mintFace = 1) Then '数值类型需要检查
        If Not CheckValid(strText, strInfo) Then Exit Function
    ElseIf mintType = 2 Then '日期类型
        If strText <> "" Then
            If InStr(1, strText, "-") = 0 Then
                If IsNumeric(strText) = False Then
                    strInfo = "日期不能包含“-”以外的字符,请检查!"
                    Exit Function
                End If
                If Len(strText) <> 8 Then
                     strInfo = "日期格式只能为[YYYY-MM-DD]或[YYYYMMDD],请检查!"
                     Exit Function
                End If
                strText = Mid(strText, 1, 4) & "-" & Mid(strText, 5, 2) & "-" & Mid(strText, 7, 2)
            Else
                If Left(strText, 1) = "-" Or Right(strText, 1) = "-" Then
                    strInfo = "日期开始和结尾不能存在“-”字符,请检查!"
                    Exit Function
                End If
            End If
            arrDate = Split(strText, "-")
            If UBound(arrDate) <> 2 Then
                strInfo = "日期格式只能为[YYYY-MM-DD]或[YYYYMMDD],请检查!"
                Exit Function
            End If
            For intIndex = 0 To UBound(arrDate)
                If IsNumeric(CStr(arrDate(intIndex))) = False Then
                    strInfo = "日期的年月日只能为数字,请检查!"
                    Exit Function
                End If
                If intIndex = 0 Then
                    If Len(CStr(arrDate(intIndex))) > 4 Then
                        strInfo = "日期年份长度不能超过4位,请检查!"
                        Exit Function
                    End If
                ElseIf intIndex = 1 Then
                    If Len(CStr(arrDate(intIndex))) > 2 Then
                        strInfo = "日期月份长度不能超过2位,请检查!"
                        Exit Function
                    End If
                Else
                    If Len(CStr(arrDate(intIndex))) > 2 Then
                        strInfo = "日期天数长度不能超过2位,请检查!"
                        Exit Function
                    End If
                End If
            Next
            If Not IsDate(Format(strText, "YYYY-MM-DD")) Then
                strInfo = "录入的日期不是有效的日期,请检查!"
                Exit Function
            End If
            strText = Format(strText, "YYYY-MM-DD")
        End If
    End If
    If strText <> strOldText Then mblnChange = True
    strReturn = strText
    CheckInput = True
End Function

Private Function CheckValid(strReturn As String, strInfo As String) As Boolean
    Dim blnCheck As Boolean
    Dim dblMin As Double, dblMax As Double
    Dim strText As String, strName As String
    '检查数据

    On Error GoTo errHand

    strName = VsfData.TextMatrix(VsfData.ROW, COL_Name)
    strText = strReturn
    mrsPartogram.Filter = 0
    mrsPartogram.Filter = "中文名='" & strName & "'"
    If strText <> "" Then
        blnCheck = True
        '如果是曲线项目,如果输入的不是数字型则不检查
        If Val(NVL(mrsPartogram!类型)) = 0 Then
            If Not IsNumeric(Trim(strText)) Then
                blnCheck = False
            End If
        End If

        If blnCheck Then
            If Val(NVL(mrsPartogram!类型, 0)) = 0 Then
                strText = Val(strText)
                If Val(NVL(mrsPartogram!小数, 0)) <> 0 Then   '长度通过控件的MaxLength来控制的
                    If InStr(1, strText, ".") <> 0 Then strText = Mid(strText, 1, InStr(1, strText, ".") - 1)
                    If Len(strText) > Val(NVL(mrsPartogram!长度)) Then
                        mrsPartogram.Filter = 0
                        strInfo = "[" & strName & "]录入的数据超过了合法精度！"
                        Exit Function
                    End If
                End If

                If Val(Replace(NVL(mrsPartogram!数值域), ";", "")) <> 0 Then
                    dblMin = Val(Split(mrsPartogram!数值域, ";")(0))
                    dblMax = Val(Split(mrsPartogram!数值域, ";")(1))
                    If Not (Val(strText) >= dblMin And Val(strText) <= dblMax) Then
                        mrsPartogram.Filter = 0
                        strInfo = "[" & strName & "]录入的数据不在" & Format(dblMin, "#0.00") & "～" & Format(dblMax, "#0.00") & "的有效范围！"
                        Exit Function
                    End If
                End If
                If Val(NVL(mrsPartogram!小数, 0)) > 0 Then
                    strText = Format(strText, "#0." & String(Val(NVL(mrsPartogram!小数, 0)), "0"))
                Else
                    strText = Format(strText, "#0")
                End If

            Else
                If LenB(StrConv(strText, vbFromUnicode)) > mrsPartogram!长度 Then
                    strInfo = "[" & strName & "]录入的数据超过了最大长度：" & mrsPartogram!长度 & "！"
                    mrsPartogram.Filter = 0
                    Exit Function
                End If
            End If
        End If
    End If

    strReturn = strText
    CheckValid = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
