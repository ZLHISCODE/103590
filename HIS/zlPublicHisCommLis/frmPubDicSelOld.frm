VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Begin VB.Form frmPubDicSelOld 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "选择器"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picItem 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5025
      Left            =   120
      ScaleHeight     =   5025
      ScaleWidth      =   4965
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   4965
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3450
         TabIndex        =   3
         Top             =   4530
         Width           =   1245
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "确定(&O)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1830
         TabIndex        =   2
         Top             =   4530
         Width           =   1245
      End
      Begin VB.TextBox txtSel 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   165
         TabIndex        =   1
         Top             =   270
         Width           =   4515
      End
      Begin VSFlex8Ctl.VSFlexGrid VSFItemSel 
         Height          =   2775
         Left            =   75
         TabIndex        =   4
         Top             =   930
         Width           =   4635
         _cx             =   8176
         _cy             =   4895
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
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
         Editable        =   2
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
End
Attribute VB_Name = "frmPubDicSelOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mRecordSource As ADODB.Recordset
Private mstrValue As String
'图片框的边框颜色
Private Const const_PicRectBackColour As Long = &HE0E0E0
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    GetRowValue
End Sub

Private Sub Form_Activate()
    Me.txtSel.SetFocus
End Sub

Private Sub Form_Resize()
    With Me.picItem
        .Top = 0
        .Left = 0
        .Width = Me.ScaleWidth - 0
        .Height = Me.ScaleHeight - 0
    End With
End Sub

Private Sub picItem_Resize()
    With Me.txtSel
        .Top = 100
        .Left = 100
        .Width = Me.picItem.ScaleWidth - 200
    End With
    
    With Me.VSFItemSel
        .Top = Me.txtSel.Top + Me.txtSel.Height + 100
        .Left = 100
        .Width = Me.picItem.ScaleWidth - 200
        .Height = Me.picItem.ScaleHeight - .Top - Me.cmdOK.Height - 300
    End With
    
    With Me.cmdCancel
        .Top = Me.VSFItemSel.Top + Me.VSFItemSel.Height + 180
        .Left = Me.ScaleWidth - .Width - 300
    End With
    
    
    With Me.cmdOK
        .Top = cmdCancel.Top
        .Left = Me.cmdCancel.Left - .Width - 300
    End With
    
    Call PicDrowBorder(picItem)
    Call PicDrowSplit(picItem, Me.VSFItemSel)

End Sub

Public Function ShowMe(formParent As Object, RecordSource As Recordset, strFind As String, Optional lngID As Long) As String
          '功能   打开公共的选择器（单列)
          '参数   RecordSource    传入要查询的记录集
          '       strField        过滤字段
          '       strFind         过滤字段的查询条件
          Dim strFilter As String
1         On Error GoTo showMe_Error

2         mstrValue = ""
3         Set mRecordSource = RecordSource
          
4         If strFind <> "" Then
5             strFilter = GetFindString(RecordSource, strFind, lngID, "")
6         Else
7             strFilter = ""
8             If lngID > 0 Then
9                 strFilter = "id=" & lngID
10            End If
11        End If
          
12        mRecordSource.Filter = strFilter
          
13        If mRecordSource.RecordCount <> 1 Then
14            If mRecordSource.RecordCount = 0 Then
15                mRecordSource.Filter = ""
16                If lngID > 0 Then
17                    strFilter = "id=" & lngID
18                End If
19                mRecordSource.Filter = strFilter
20            End If
21            Load frmPubDicSelOld
22            InitPublicDicVsf Me.VSFItemSel, mRecordSource, ""
23            frmPubDicSelOld.Show vbModal, formParent
24            If mRecordSource.RecordCount > 0 Then
25                Me.txtSel.Text = strFind
26            End If
27        Else
28            InitPublicDicVsf Me.VSFItemSel, mRecordSource, ""
29            mstrValue = GetVSFRowValue(VSFItemSel, VSFItemSel.Row, "")
30        End If
31        ShowMe = mstrValue


32        Exit Function
showMe_Error:
33        Call WriteErrLog("zlPublicHisCommLis", "frmPubDicSelOld", "执行(showMe)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
34        Err.Clear
End Function

Private Sub txtSel_Change()
    If txtSel.Text <> "" Then
        mRecordSource.Filter = GetFindString(mRecordSource, txtSel.Text, , "")
    Else
        mRecordSource.Filter = ""
    End If
    InitPublicDicVsf Me.VSFItemSel, mRecordSource, ""
End Sub

Private Sub txtSel_GotFocus()
    Call TextSelAll(txtSel)
End Sub

Private Sub txtSel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then
        '向上按键
        With Me.VSFItemSel
            If .Row > 1 Then
                .Row = .Row - 1
            End If
        End With
    End If
    If KeyCode = 40 Then
        With Me.VSFItemSel
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
            End If
        End With
    End If
    
End Sub

Private Sub txtSel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        GetRowValue
    End If
End Sub

Private Sub VSFItemSel_Click()
    Me.txtSel.SetFocus
End Sub

Private Sub VSFItemSel_DblClick()
    GetRowValue
End Sub

Private Sub VSFItemSel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        GetRowValue
    End If
End Sub
Private Function GetRowValue()
    '功能           返回行结果
    mstrValue = GetVSFRowValue(VSFItemSel, VSFItemSel.Row, "")
    Unload Me
End Function
Public Function InitPublicDicVsf(VSFList As VSFlexGrid, RecordSource As Recordset, ByRef strErr As String) As Boolean
    
    On Error GoTo errH
    '初使化表格控件
    If Not vfgLoadFromRecord(VSFList, RecordSource, strErr) Then Exit Function
    
    With VSFList
        .ColWidth(1) = 1300: .ColHidden(1) = False
        .ColWidth(2) = 2500: .ColHidden(2) = False
        .ColWidth(3) = 300: .ColHidden(3) = False
    End With
errH:
    strErr = Err.Number & " " & Err.Description
End Function

Public Function GetFindString(RecordSource As Recordset, strFind As String, Optional lngID As Long, Optional ByRef strErr As String) As String
          '功能   从数据源中提取过滤字段并生成过滤字串
          '参数   RecordSource 数据源
          '       strFind 过滤字串
          
          Dim intloop As Integer
1         On Error GoTo GetFindString_Error

2         For intloop = 1 To RecordSource.Fields.Count - 1
3             If RecordSource.Fields(intloop).Type = 200 Then
4                 If lngID = 0 Then
5                     GetFindString = GetFindString & "or " & RecordSource.Fields(intloop).Name & " like '*" & StringDelInvalidWord(strFind) & "*' "
6                 Else
7                     GetFindString = GetFindString & "or (" & RecordSource.Fields(intloop).Name & " like '*" & StringDelInvalidWord(strFind) & "*' " & _
                                      " and id = " & lngID & " )"
8                 End If
9             End If
10        Next
11        If GetFindString <> "" Then
12            GetFindString = Mid(GetFindString, 3)
13        End If


14        Exit Function
GetFindString_Error:
15        Call WriteErrLog("zlPublicHisCommLis", "frmPubDicSelOld", "执行(GetFindString)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
16        Err.Clear

End Function

Public Function GetVSFRowValue(VSFList As VSFlexGrid, intRow As Integer, ByRef strErr As String) As String
          '功能       到得当前行的值
          Dim intloop As Integer
1         On Error GoTo GetVSFRowValue_Error

2         With VSFList
3             For intloop = 0 To VSFList.Cols - 1
4                 GetVSFRowValue = GetVSFRowValue & "," & .TextMatrix(intRow, intloop)
5             Next
6             GetVSFRowValue = Mid(GetVSFRowValue, 2)
7         End With


8         Exit Function
GetVSFRowValue_Error:
9         Call WriteErrLog("zlPublicHisCommLis", "frmPubDicSelOld", "执行(GetVSFRowValue)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
10        Err.Clear

End Function
Public Sub PicDrowBorder(Picobj As PictureBox, Optional lngLineColour As Long = -1)
    '功能       画图片边框
    On Error Resume Next
    With Picobj
        .AutoRedraw = True
        .Cls
        .DrawWidth = 2
        
        If lngLineColour = -1 Then
            .ForeColor = const_PicRectBackColour
        Else
            .ForeColor = lngLineColour
        End If
        Picobj.Line (25, 25)-(.Width - 50, .Height - 50), , B
    End With

End Sub
Public Sub PicDrowSplit(Picobj As PictureBox, objSplit As Object, Optional lngHeightSplit As Long)
    '功能       画图片的分隔线
    On Error Resume Next
    With Picobj
        .AutoRedraw = True
'        .ForeColor = const_PicRectBackColour
        If lngHeightSplit = 0 Then
            Picobj.Line (25, objSplit.Top + objSplit.Height + 70)-(.Width - 50, objSplit.Top + objSplit.Height + 70), , B
        Else
            Picobj.Line (25, objSplit.Top + objSplit.Height + lngHeightSplit)-(.Width - 50, objSplit.Top + objSplit.Height + lngHeightSplit), , B
        End If
    End With

End Sub
Public Sub TextSelAll(objText As TextBox)
    objText.SelStart = 0
    objText.SelLength = Len(objText)
End Sub

Public Function vfgLoadFromRecord(ByRef objVfg As VSFlexGrid, _
                                  ByRef rsTmp As ADODB.Recordset, _
                                  ByRef strErr As String, _
                                  Optional objImgList As ImageList) As Boolean
          '将记录集数据装入vfg控件
          'objVfg : vfg控件
          'rsTmp  : 装入控件的记录集
          'strErr :提示信息
          Dim i As Integer, strTitle As String
          
          '标题
1         On Error GoTo vfgLoadFromRecord_Error

2         For i = 0 To rsTmp.Fields.Count - 1
3             strTitle = strTitle & ";" & rsTmp.Fields(i).Name & ",0," & flexAlignLeftCenter
4         Next
5         If strTitle <> "" Then strTitle = Mid(strTitle, 2)
          
6         Call vfgSetting(0, objVfg, strTitle, objImgList)
          
          '处理数据
7         With objVfg
8             .Tag = "A"
9             .Rows = .FixedRows + 1
10            .Cell(flexcpText, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = ""
              'Set .DataSource = rsTmp 直接设数据源，则原来设置的格式标题等格式丢失，需手工添加数据
11            Do Until rsTmp.EOF
12                For i = 0 To rsTmp.Fields.Count - 1
13                    .TextMatrix(.Rows - 1, i) = CStr("" & rsTmp.Fields(i).value)
14                    If Not objImgList Is Nothing Then
15                        If CheckImgListKey(objImgList, rsTmp.Fields(i).Name) = True And CheckImgListKey(objImgList, rsTmp.Fields(i).value & "") = True Then
16                            .Row = .Rows - 1
17                            .Col = i
18                            .CellPicture = objImgList.ListImages(rsTmp.Fields(i).value).ExtractIcon
19                        End If
20                    End If
21                Next
22                .Rows = .Rows + 1
23                rsTmp.MoveNext
24            Loop
25            If .Rows > .FixedRows + 1 Then .Rows = .Rows - 1
26            .Tag = ""
27        End With
28        vfgLoadFromRecord = True


29        Exit Function
vfgLoadFromRecord_Error:
30        Call WriteErrLog("zlPublicHisCommLis", "frmPubDicSelOld", "执行(vfgLoadFromRecord)时发生错误,错误号:" & Err.Number & " 出错原因:" & Err.Description & " 错误行：" & Erl, True)
31        Err.Clear

End Function

Public Sub vfgSetting(ByVal LngStyle As Long, ByRef objVfg As VSFlexGrid, Optional ByVal strTtile As String, Optional VsfImg As ImageList)
    'lngStyle＝0 默认设置，统一Vfg表格的外观
    'strHead：  标题格式串
    '           标题1,宽度,对齐方式;标题2,宽度,对齐方式;.......
    '           对齐方式取值, * 表示常用取值
    '           FlexAlignLeftTop       0   左上
    '           flexAlignLeftCenter    1   左中  *
    '           flexAlignLeftBottom    2   左下
    '           flexAlignCenterTop     3   中上
    '           flexAlignCenterCenter  4   居中  *
    '           flexAlignCenterBottom  5   中下
    '           flexAlignRightTop      6   右上
    '           flexAlignRightCenter   7   右中  *
    '           flexAlignRightBottom   8   右下
    '           flexAlignGeneral       9   常规
    'objVfg:    要初始化的控件
    'VsfImg:    ImageList图标集控件对象

    Dim arrHead As Variant, i As Long, strHead As String
    If strTtile = "" Then
        strHead = "第1列,900,1;第2列,900,1;第3列,900,1"
    Else
        strHead = strTtile
    End If
    arrHead = Split(strHead, ";")
    
    
    With objVfg
        .Tag = "A"
        '1.边框
        .Appearance = flex3DLight
        .BorderStyle = flexBorderFlat
        .GridLines = flexGridFlat
        .GridColorFixed = flexGridFlat
        
        '2.颜色
        .BackColor = vbWindowBackground '窗口背景
        .BackColorAlternate = vbWindowBackground
        .BackColorBkg = vbWindowBackground
        .BackColorFixed = vbButtonFace '按钮表面
        .BackColorFrozen = &H0&         '黑
        .FloodColor = &HC0&             '红
        .BackColorSel = &HFFEBD7        '浅绿
        .ForeColor = vbWindowText       '窗口文本
        .ForeColorFixed = vbButtonText  '按钮文本
        .ForeColorFrozen = &H0&         '黑
        .ForeColorSel = vbWindowText
        
        .GridColor = vbApplicationWorkspace '应用程序工作区
        .GridColorFixed = vbApplicationWorkspace
        .SheetBorder = vbWindowBackground
        .TreeColor = vbButtonShadow         '按钮阴影
        
        '3.初始化行列

        .Redraw = False
        .Clear
        .Cols = 2
        .FixedRows = 1: .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        
        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .ColKey(i) = Split(arrHead(i), ",")(0) '将标提作为colKey值
            If CheckImgListKey(VsfImg, .TextMatrix(.FixedRows - 1, .FixedCols + i)) = True Then
                .Row = .FixedRows - 1
                .Col = .FixedCols + i
                .CellPicture = VsfImg.ListImages(Split(arrHead(i), ",")(0)).ExtractIcon
                '有图标时不显示文字
                .TextMatrix(.FixedRows - 1, .FixedCols + i) = ""
            End If
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                .ColWidth(.FixedCols + i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                
                '为了支持zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0 '为了支持zl9PrintMode
            End If
        Next
        
        '固定行文字居中
        If .FixedRows > 0 And .Cols > 0 Then
            .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        End If
        .RowHeight(0) = 300
        .RowHeightMin = 300
        .WordWrap = True '自动换行
        .AutoSizeMode = flexAutoSizeRowHeight '自动行高
        .AutoResize = True '自动
        .Redraw = True
        
        
        '4.其他属性
        .SelectionMode = flexSelectionByRow     '整行选择
        .ExplorerBar = flexExNone               '点标题栏不响应（排序及移动列）操作
        .AllowUserResizing = flexResizeColumns  '可调整列宽
        .Editable = flexEDNone                  '只读
        .Tag = ""
    End With
    
End Sub
Public Function CheckImgListKey(Vfgimg As ImageList, strKey As String) As Boolean
    '功能           检查关键字是否在图像列表中存在，如果存在返回为真
    '参数
    '               Vfgimg 传入的图像对象
    '               strKey 要检查当前传入的Key是否存在
    '返回           有返回真，没有返回假
    Dim intloop As Integer
    On Error Resume Next
    If Vfgimg Is Nothing Then Exit Function
    With Vfgimg
        For intloop = 1 To .ListImages.Count
            If .ListImages(intloop).Key = strKey Then
                CheckImgListKey = True
                Exit Function
            End If
        Next
    End With
End Function
