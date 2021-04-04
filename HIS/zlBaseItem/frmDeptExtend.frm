VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDeptExtend 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "部门扩展信息"
   ClientHeight    =   6600
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   10845
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   10845
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      Height          =   1700
      Index           =   0
      Left            =   120
      ScaleHeight     =   1629.323
      ScaleMode       =   0  'User
      ScaleWidth      =   2235
      TabIndex        =   3
      Top             =   4185
      Visible         =   0   'False
      Width           =   2300
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   350
      Left            =   7275
      TabIndex        =   2
      Top             =   5670
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8445
      TabIndex        =   1
      Top             =   5670
      Width           =   1100
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   3765
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9630
      _cx             =   16986
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483643
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   3
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   255
      RowHeightMax    =   3000
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDeptExtend.frx":0000
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
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      DataMode        =   0
      VirtualData     =   0   'False
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
   Begin MSComDlg.CommonDialog cdl照片 
      Left            =   7845
      Top             =   4740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image img 
      Height          =   1700
      Index           =   0
      Left            =   2505
      Stretch         =   -1  'True
      Top             =   4245
      Visible         =   0   'False
      Width           =   2300
   End
End
Attribute VB_Name = "frmDeptExtend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlngId As Long
Private mblnPro As Boolean '是否有扩展项目
Private mintType As Integer '1-人员；0-部门
Private mblnEdit As Boolean '内容是否被修改
Private mbln图片 As Boolean '是否有图片
Private mbln图片更改 As Boolean '是否更改了图片
Private mintIndex As Integer '图片容器索引
Private mint编辑状态 As Integer '1-可编辑，0-不可编辑
Private mintCountPic As Integer '图片个数
Private mstrName As String   '窗体名称

Public Sub ShowMe(ByVal fraPar As Form, ByVal strID As String, ByVal strName As String, Optional ByVal intType As Integer, Optional int编辑状态 As Integer)
    mlngId = Val(strID)
    mintType = intType
    mint编辑状态 = int编辑状态
    mstrName = strName
    
    If mintType = 1 Then
        Me.Caption = "人员扩展信息-" & mstrName
    Else
        Me.Caption = "部门扩展信息-" & mstrName
    End If
    
    Call initVSf(mlngId, mintType)
    
    If mblnPro Then Me.Show vbModal, fraPar
End Sub

Public Sub ReadBlob(ByVal lngId As Long, ByVal strName As String, ByVal intIndex As Integer, ByVal intType As Integer)
    '读取图片
    Dim strTempFile As String
    
    '初始化图片位置尺寸
    img(intIndex).Left = pic(intIndex).ScaleLeft
    img(intIndex).Top = pic(intIndex).ScaleTop
    img(intIndex).Width = pic(intIndex).ScaleWidth
    img(intIndex).Height = pic(intIndex).ScaleHeight
    
    If intType = 1 Then '人员
        strTempFile = sys.Readlob(100, 20, lngId & "," & strName)
    Else
        strTempFile = sys.Readlob(100, 19, lngId & "," & strName)
    End If
    
    img(intIndex).Tag = ""
    img(intIndex).Picture = Nothing
    pic(intIndex).Picture = Nothing
    pic(intIndex).AutoRedraw = True
    
    '处理图片
    If strTempFile <> "" Then
        img(intIndex).Tag = strTempFile
        img(intIndex).Picture = LoadPicture(strTempFile)
        pic(intIndex).PaintPicture img(intIndex).Picture, 0, 0, pic(intIndex).Width, pic(intIndex).Height
        
    End If
End Sub

Private Function SaveBlob(ByVal lngId As Long, ByVal strName As String, ByVal intIndex As Integer) As Boolean
    '保存图片
    Dim blnOk As Boolean
    
    On Error GoTo ErrHandle
    
    If img(intIndex).Tag = "" Then
        If mintType = 1 Then '人员
            gstrSQL = "Update 人员扩展信息 Set 图片=Null Where 人员id=" & lngId & " And 项目='" & strName & "'"
        Else
            gstrSQL = "Update 部门扩展信息 Set 图片=Null Where 部门id=" & lngId & " And 项目='" & strName & "'"
        End If
        gcnOracle.Execute gstrSQL
        blnOk = True
    Else
        If mintType = 1 Then '人员
            blnOk = sys.Savelob(100, 20, lngId & "," & strName, img(intIndex).Tag)
        Else
            blnOk = sys.Savelob(100, 19, lngId & "," & strName, img(intIndex).Tag)
        End If
    End If
    
    SaveBlob = blnOk
    
    Exit Function
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Public Sub initVSf(ByVal lngId As Long, Optional ByVal intType As Integer)
    '初始化vsf表格，添加数据
    Dim rsTemp As ADODB.Recordset
    Dim rs信息 As ADODB.Recordset
    Dim intIndex As Integer
    Dim intRow As Integer
    Dim i As Integer
    Dim bln图片 As Boolean
    
    On Error GoTo ErrHandle
    
    If intType = 1 Then '人员
        gstrSQL = "Select 编码, 名称, Nvl(是否图片, 0) As 是否图片 From 人员扩展项目 Order By 编码"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "人员扩展属性")
    Else
        gstrSQL = "Select 编码, 名称, Nvl(是否图片, 0) As 是否图片 From 部门扩展项目 Order By 编码"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "部门扩展属性")
    End If
    
    With VSFList
        .Rows = 1
        .MergeCells = flexMergeFree
        .MergeRow(0) = True
        .RowHeightMin = 255
        .RowHeightMax = 2000
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        For i = 1 To img.Count - 1
            Unload img(i)
        Next
        For i = 1 To pic.Count - 1
            Unload pic(i)
        Next
        
        If rsTemp.RecordCount > 0 Then
            .redraw = flexRDNone
            Do While Not rsTemp.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, .ColIndex("编码")) = rsTemp!编码
                .TextMatrix(.Rows - 1, .ColIndex("项目")) = rsTemp!名称
                
                '内容
                If intType = 1 Then '人员
                    gstrSQL = "Select 内容 From 人员扩展信息 Where 人员id=[1] And 项目=[2]"
                    Set rs信息 = zlDatabase.OpenSQLRecord(gstrSQL, "人员扩展属性", lngId, .TextMatrix(.Rows - 1, .ColIndex("项目")))
                Else
                    gstrSQL = "Select 内容 From 部门扩展信息 Where 部门id=[1] And 项目=[2]"
                    Set rs信息 = zlDatabase.OpenSQLRecord(gstrSQL, "部门扩展属性", lngId, .TextMatrix(.Rows - 1, .ColIndex("项目")))
                End If
                
                If Not rs信息.EOF Then
                    .TextMatrix(.Rows - 1, .ColIndex("内容")) = IIF(IsNull(rs信息!内容), "", rs信息!内容)
                End If
                
                '图片
                If rsTemp!是否图片 = 1 Then
                    bln图片 = True
                    If mint编辑状态 = 0 Then
                        intIndex = 0
                    Else
                        If intIndex <> 0 Then
                            Load img(intIndex)
                            Load pic(intIndex)
                        End If
                    End If
                    
                    Call ReadBlob(lngId, .TextMatrix(.Rows - 1, .ColIndex("项目")), intIndex, intType)
                    
                    .Cell(flexcpPicture, .Rows - 1, .ColIndex("图片"), .Rows - 1, .ColIndex("图片")) = pic(intIndex).Image
                    .RowHeight(.Rows - 1) = 1700
                    .TextMatrix(.Rows - 1, .ColIndex("选择图片")) = "…"
                    .TextMatrix(.Rows - 1, .ColIndex("清除图片")) = "×"
                    intIndex = intIndex + 1
                Else
                    If .TextMatrix(.Rows - 1, .ColIndex("内容")) = "" Then
                        .TextMatrix(.Rows - 1, .ColIndex("图片")) = " "
                        .TextMatrix(.Rows - 1, .ColIndex("内容")) = " "
                    Else
                        .TextMatrix(.Rows - 1, .ColIndex("图片")) = .TextMatrix(.Rows - 1, .ColIndex("内容"))
                    End If
                    .MergeRow(.Rows - 1) = True
                    .RowHeight(.Rows - 1) = 1000
                End If
                
                rsTemp.MoveNext
            Loop
            
            If Not bln图片 Then
                .ColHidden(.ColIndex("选择图片")) = True
                .ColHidden(.ColIndex("清除图片")) = True
            End If
            
            .redraw = flexRDDirect
            mblnPro = True
            Call FS.ShowTipInfo(VSFList.hwnd, "")
        Else
            If mint编辑状态 = 0 Then Exit Sub
            If intType = 1 Then '人员
                MsgBox "未设置人员扩展项目，请到数据字典->人员属性->人员扩展项目中设置！", vbInformation, gstrSysName
            Else
                MsgBox "未设置部门扩展项目，请到数据字典->部门属性->部门扩展项目中设置！", vbInformation, gstrSysName
            End If
            mblnPro = False
        End If
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If SaveData Then
        MsgBox "保存成功！", vbInformation, gstrSysName
        mblnEdit = False
        mbln图片更改 = False
        Unload Me
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    If mint编辑状态 = 0 Then
        cmdCancel.Visible = False
        cmdOK.Visible = False
        VSFList.ColHidden(VSFList.ColIndex("选择图片")) = True
        VSFList.ColHidden(VSFList.ColIndex("清除图片")) = True
        VSFList.Move 10, 10, Me.ScaleWidth - 20, Me.ScaleHeight - 20
        VSFList.ColWidth(VSFList.ColIndex("内容")) = VSFList.Width - VSFList.ColWidth(VSFList.ColIndex("项目")) - VSFList.ColWidth(VSFList.ColIndex("图片")) - 400
        Exit Sub
    End If
    
    cmdCancel.Move Me.Width - cmdCancel.Width - 300, Me.ScaleHeight - cmdCancel.Height - 50
    cmdOK.Move cmdCancel.Left - cmdOK.Width - 10, cmdCancel.Top
    VSFList.Move 10, 10, Me.ScaleWidth - 20, Me.ScaleHeight - cmdOK.Height - 100
    VSFList.ColWidth(VSFList.ColIndex("内容")) = VSFList.Width - VSFList.ColWidth(VSFList.ColIndex("项目")) - VSFList.ColWidth(VSFList.ColIndex("图片")) - VSFList.ColWidth(VSFList.ColIndex("选择图片")) - VSFList.ColWidth(VSFList.ColIndex("清除图片")) - 400
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnEdit Or mbln图片更改 Then
        If MsgBox("已修改内容还未保存，是否确定退出？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = 1
        Else
            mblnEdit = False
            mbln图片更改 = False
        End If
    End If
End Sub

Private Sub vsfList_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With VSFList
        If .TextMatrix(.Row, .ColIndex("选择图片")) = "" Then
            If Trim(.TextMatrix(Row, .ColIndex("图片"))) = "" Then
                .TextMatrix(Row, .ColIndex("图片")) = " "
                .TextMatrix(Row, .ColIndex("内容")) = " "
            Else
                .TextMatrix(Row, .ColIndex("内容")) = .TextMatrix(Row, .ColIndex("图片"))
            End If
        End If
    End With
End Sub

Private Sub vsfList_ChangeEdit()
    mblnEdit = True
End Sub

Private Sub vsfList_EnterCell()
    If mint编辑状态 = 0 Then Exit Sub
    With VSFList
        If .TextMatrix(.Row, .ColIndex("选择图片")) = "" Then
            If .Col = .ColIndex("图片") Then
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
            End If
        Else
            If .Col = .ColIndex("内容") Then
                .Editable = flexEDKbdMouse
            Else
                .Editable = flexEDNone
            End If
        End If
    End With
End Sub

Private Sub VSFList_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    If KeyAscii = vbKeyDelete Or KeyAscii = vbKeyBack Then Exit Sub
    With VSFList
        If LenB(StrConv(.EditText + Chr(KeyAscii), vbFromUnicode)) > 1000 Then
            KeyAscii = 0
        End If
    End With
    
End Sub

Private Sub vsfList_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Integer
    Dim dblHeight As Double
    Dim dblWidth As Double
    
    With VSFList
        If .Rows = 1 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("选择图片")) = "" Then
            If .Col = .ColIndex("图片") Then
                Call FS.ShowTipInfo(VSFList.hwnd, Trim(.TextMatrix(.Row, .ColIndex("图片"))), True)
            Else
                Call FS.ShowTipInfo(VSFList.hwnd, "")
            End If
        Else
            If .Col = .ColIndex("内容") Then
                Call FS.ShowTipInfo(VSFList.hwnd, Trim(.TextMatrix(.Row, .ColIndex("内容"))), True)
            Else
                Call FS.ShowTipInfo(VSFList.hwnd, "")
            End If
            
            If .Col = .ColIndex("选择图片") Or .Col = .ColIndex("清除图片") Then
                For i = 0 To .Rows - 1
                    dblHeight = dblHeight + .RowHeight(i)
                Next
                
                For i = 0 To .Cols - 1
                    If .ColHidden(i) = False Then
                        dblWidth = dblWidth + .ColWidth(i)
                    End If
                Next
                
                If X < dblWidth And Y > .RowHeight(0) And Y < dblHeight Then
                    If .Col = .ColIndex("选择图片") Then
                        Call SelectPic
                    Else
                        Call ClearPic
                    End If
                End If
            End If
        End If
    End With
End Sub

Private Sub SelectPic()
    '选择图片
    Dim intIndex As Integer
    
    With cdl照片
        .CancelError = True
        .Filter = "图片文件(*.bmp,*.gif,*.jpg)|*.bmp;*.gif;*.jpg"
        
        On Error Resume Next
        .ShowOpen
        intIndex = GetPicStaition(VSFList.Row)
        If Err <> 0 Then
            '没选中文件
            Err.Clear
        Else
            img(intIndex).Picture = LoadPicture(.FileName)
            pic(intIndex).PaintPicture img(intIndex).Picture, 0, 0, pic(intIndex).Width, pic(intIndex).Height
            VSFList.Cell(flexcpPicture, VSFList.Row, VSFList.ColIndex("图片"), VSFList.Row, VSFList.ColIndex("图片")) = pic(intIndex).Image
            
            If Err <> 0 Then
                MsgBox "图片文件无效，或文件不存在！", vbInformation, gstrSysName
                Exit Sub
            End If
            img(intIndex).Tag = .FileName
            mbln图片 = True
            mbln图片更改 = True
        End If
    End With
End Sub

Private Function GetPicStaition(ByVal intCRow As Integer) As Integer
    '获取加载图片索引
    Dim intRow As Integer
    
    With VSFList
        mintCountPic = -1
        For intRow = 1 To intCRow
            If .TextMatrix(intRow, .ColIndex("选择图片")) = "…" Then
                mintCountPic = mintCountPic + 1
            End If
        Next
        GetPicStaition = mintCountPic
    End With
End Function

Private Sub ClearPic()
    '清除图片
    Dim intIndex As Integer
    
    intIndex = GetPicStaition(VSFList.Row)
    
    If img(intIndex).Tag = "" Then Exit Sub
    
    mbln图片 = False
    mbln图片更改 = True
    img(intIndex).Tag = ""
    img(intIndex).Picture = Nothing
    pic(intIndex).Picture = Nothing
    VSFList.Cell(flexcpPicture, VSFList.Row, VSFList.ColIndex("图片"), VSFList.Row, VSFList.ColIndex("图片")) = pic(intIndex).Image
End Sub

Private Function SaveData() As Boolean
    '保存数据
    Dim blnTran As Boolean
    Dim intRow As Integer
    Dim arrSQL As Variant
    
    On Error GoTo ErrHandle
    
    SaveData = False
    arrSQL = Array()
    
    With VSFList
        For intRow = 1 To .Rows - 1
            If Check项目(.TextMatrix(intRow, .ColIndex("项目"))) Then
                If LenB(StrConv(.TextMatrix(intRow, .ColIndex("内容")), vbFromUnicode)) > 1000 Then
                    MsgBox "第" & intRow & "行“" & .TextMatrix(intRow, .ColIndex("项目")) & "”输入内容大于1000个字符，请重新输入！", vbInformation, gstrSysName
                    .Col = .Col
                    .Row = intRow
                    .SetFocus
                    Exit Function
                End If
                If mintType = 1 Then '人员
                    gstrSQL = "Zl_人员扩展信息_Delete(" & mlngId & ",'" & .TextMatrix(intRow, .ColIndex("项目")) & "')"

                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = gstrSQL
                    
                    gstrSQL = "Zl_人员扩展信息_Insert(" & mlngId & ",'" & .TextMatrix(intRow, .ColIndex("项目")) & "','" & .TextMatrix(intRow, .ColIndex("内容")) & "')"

                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = gstrSQL
                Else
                    gstrSQL = "Zl_部门扩展信息_Delete(" & mlngId & ",'" & .TextMatrix(intRow, .ColIndex("项目")) & "')"

                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = gstrSQL
                    
                    gstrSQL = "Zl_部门扩展信息_Insert(" & mlngId & ",'" & .TextMatrix(intRow, .ColIndex("项目")) & "','" & .TextMatrix(intRow, .ColIndex("内容")) & "')"

                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = gstrSQL
                End If
            End If
        Next
    
        gcnOracle.BeginTrans: blnTran = True
        For intRow = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(intRow)), "SaveData")
        Next
        
        For intRow = 1 To .Rows - 1
            If Check项目(.TextMatrix(intRow, .ColIndex("项目"))) Then
                If .TextMatrix(intRow, .ColIndex("选择图片")) <> "" Then
                    Call SaveBlob(mlngId, .TextMatrix(intRow, .ColIndex("项目")), GetPicStaition(intRow))
                End If
            End If
        Next
        gcnOracle.CommitTrans: blnTran = False
    End With
    
    SaveData = True
    
    Exit Function
ErrHandle:
    If blnTran = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Check项目(ByVal strName As String) As Boolean
    Dim rsTemp As Recordset
    
    Check项目 = False
    If mintType = 1 Then '人员
        gstrSQL = "Select 1 From 人员扩展项目 Where 名称 = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "人员扩展属性", strName)
    Else
        gstrSQL = "Select 1 From 部门扩展项目 Where 名称 = [1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "部门扩展属性", strName)
    End If
    
    If Not rsTemp.EOF Then Check项目 = True
    
End Function
